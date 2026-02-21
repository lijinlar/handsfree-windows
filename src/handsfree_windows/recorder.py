"""Passive macro recorder for handsfree-windows.

Records user mouse clicks and keyboard input in the background, resolves UIA
selectors for each interaction, and saves a replay-compatible YAML macro.

Usage (via CLI):
    hf record --out macro.yaml --passive

Stop recording by pressing F9.
"""
from __future__ import annotations

import threading
import time
from pathlib import Path
from typing import Any

import yaml

from . import uia

_IDLE_FLUSH_SECS = 1.5  # flush type buffer after this many seconds of inactivity


def passive_record(out: Path, verbose: bool = False) -> None:
    """Record user actions passively until F9 is pressed.

    - Left-click  → resolve UIA element → record `click` step
    - Printable keys → accumulate into type buffer
    - Enter key   → flush type buffer with ``enter: true``
    - Non-printable keys (arrows, backspace …) → flush type buffer, ignore key
    - Idle 1.5 s  → flush type buffer automatically
    - F9          → stop recording, flush buffer, save YAML

    Args:
        out: Path to write the YAML macro file.
        verbose: If True, print each recorded step to stdout.
    """
    steps: list[dict[str, Any]] = []

    # Shared mutable state (all access under _lock)
    _lock = threading.RLock()
    _state: dict[str, Any] = {
        "type_buffer": [],          # accumulated chars not yet flushed
        "last_selector": None,      # selector from the most recent click
        "last_type_time": 0.0,      # monotonic timestamp of last keystroke
        "last_step_time": 0.0,      # monotonic timestamp of the last recorded step
    }

    _stop = threading.Event()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _take_delay_ms() -> int:
        """Return ms since the last recorded step and reset the timer (must be called with _lock held)."""
        now = time.monotonic()
        last = _state["last_step_time"]
        delay = int((now - last) * 1000) if last > 0 else 0
        _state["last_step_time"] = now
        return delay

    def _flush_type(enter: bool = False) -> None:
        """Flush the type buffer as a `type` step (must be called with _lock held)."""
        buf: list[str] = _state["type_buffer"]
        sel = _state["last_selector"]
        if not buf and not enter:
            return
        text = "".join(buf)
        if sel is not None or text:
            step: dict[str, Any] = {
                "action": "type",
                "args": {
                    "selector_candidates": [sel] if sel else [],
                    "text": text,
                    "enter": enter,
                    "timeout": 20,
                    "delay_before": _take_delay_ms(),
                },
            }
            steps.append(step)
            if verbose:
                tag = "[type+↵]" if enter else "[type]"
                print(f"  {tag} {repr(text)}")
        _state["type_buffer"] = []
        _state["last_type_time"] = 0.0

    def _flush_safe(enter: bool = False) -> None:
        """Thread-safe wrapper around _flush_type."""
        with _lock:
            _flush_type(enter=enter)

    # ------------------------------------------------------------------
    # Mouse listener
    # ------------------------------------------------------------------

    def on_click(x: int, y: int, button: Any, pressed: bool) -> None:
        try:
            from pynput.mouse import Button  # type: ignore
        except ImportError:
            return

        if button != Button.left or not pressed:
            return

        # Flush pending typing (element focus changed)
        _flush_safe()

        # Resolve UIA element at click position
        sel = None
        try:
            elem = uia.element_from_point(x, y)
            sel = uia.selector_for_element(elem)
        except Exception as exc:
            if verbose:
                print(f"  [click] UIA lookup failed at ({x},{y}) — recording coords only: {exc}")

        with _lock:
            _state["last_selector"] = sel  # may be None for system UI
            step_args: dict = {
                "x": x,
                "y": y,
                "timeout": 20,
                "delay_before": _take_delay_ms(),
            }
            if sel is not None:
                step_args["selector_candidates"] = [sel]

            steps.append({"action": "click", "args": step_args})

        if verbose:
            if sel is not None:
                win_title = (sel.get("window") or {}).get("title", "")
                targets = sel.get("targets") or []
                ctrl_name = (targets[0] if targets else {}).get("name", "")
                print(f"  [click] window={repr(win_title)}  ctrl={repr(ctrl_name)}  ({x},{y})")
            else:
                print(f"  [click] coords-only ({x},{y})")

    # ------------------------------------------------------------------
    # Keyboard listener
    # ------------------------------------------------------------------

    def on_key_press(key: Any) -> bool | None:
        try:
            from pynput.keyboard import Key, KeyCode  # type: ignore
        except ImportError:
            return None

        # Stop hotkey
        if key == Key.f9:
            _stop.set()
            return False  # signals pynput to stop the listener

        # Enter → flush with enter=True
        if key == Key.enter:
            _flush_safe(enter=True)
            return None

        # Printable character → accumulate
        if isinstance(key, KeyCode) and key.char is not None:
            with _lock:
                _state["type_buffer"].append(key.char)
                _state["last_type_time"] = time.monotonic()
            return None

        # Any other special key → flush buffer (context change), ignore key itself
        _flush_safe()
        return None

    # ------------------------------------------------------------------
    # Idle flush thread
    # ------------------------------------------------------------------

    def _idle_flusher() -> None:
        while not _stop.is_set():
            time.sleep(0.25)
            now = time.monotonic()
            with _lock:
                last_t = _state["last_type_time"]
                has_buf = bool(_state["type_buffer"])
                timed_out = last_t > 0 and (now - last_t) > _IDLE_FLUSH_SECS
            if has_buf and timed_out:
                _flush_safe()

    idle_thread = threading.Thread(target=_idle_flusher, daemon=True)
    idle_thread.start()

    # ------------------------------------------------------------------
    # Start listeners
    # ------------------------------------------------------------------
    try:
        from pynput import keyboard, mouse  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pynput is required for passive recording. Install it with: pip install pynput"
        ) from exc

    mouse_listener = mouse.Listener(on_click=on_click)
    kb_listener = keyboard.Listener(on_press=on_key_press)

    mouse_listener.start()
    kb_listener.start()

    print("🔴 Passive recording started. Use the app normally. Press F9 to stop.")

    _stop.wait()

    # Tear down listeners
    mouse_listener.stop()
    try:
        kb_listener.stop()
    except Exception:
        pass

    # Final flush of any remaining type buffer
    _flush_safe()

    # Save macro
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        yaml.safe_dump(steps, sort_keys=False, allow_unicode=True),
        encoding="utf-8",
    )

    print(f"\n✅ Recording stopped. {len(steps)} step(s) saved → {out}")
