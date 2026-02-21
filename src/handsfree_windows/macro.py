from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from . import uia, wininput


@dataclass
class MacroStep:
    action: str
    args: dict[str, Any]


def load_macro(path: str | Path) -> list[MacroStep]:
    p = Path(path)
    data = yaml.safe_load(p.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("Macro YAML must be a list of steps")
    steps: list[MacroStep] = []
    for i, item in enumerate(data):
        if not isinstance(item, dict) or "action" not in item:
            raise ValueError(f"Invalid step at index {i}: expected mapping with 'action'")
        action = str(item["action"])
        args = dict(item.get("args", {}) or {})
        steps.append(MacroStep(action=action, args=args))
    return steps


def run_macro(path: str | Path) -> None:
    steps = load_macro(path)
    current_window = None

    _MAX_DELAY_MS = 5000  # cap replay delays at 5 s (avoids freezing on long recording pauses)

    for step in steps:
        a = step.action
        args = step.args

        # Honour recorded inter-step timing (capped at _MAX_DELAY_MS)
        delay_before = args.get("delay_before", 0)
        if delay_before and delay_before > 0:
            time.sleep(min(int(delay_before), _MAX_DELAY_MS) / 1000.0)

        if a == "focus":
            current_window = uia.focus_window(
                title=args.get("title"),
                title_regex=args.get("title_regex"),
                handle=args.get("handle"),
            )

        elif a == "start":
            # Start menu launch (human-style)
            from pywinauto.keyboard import send_keys

            app_name = str(args.get("app"))
            delay_ms = int(args.get("delay_ms", 250))

            send_keys("{VK_LWIN}")
            time.sleep(max(0, delay_ms) / 1000.0)
            send_keys(app_name, with_spaces=True)
            time.sleep(0.1)
            send_keys("{ENTER}")

        elif a == "click":
            fallback_x = args.get("x")
            fallback_y = args.get("y")
            has_selectors = bool(args.get("selector_candidates") or args.get("selector"))

            if not has_selectors and fallback_x is not None and fallback_y is not None:
                # Coord-only step (recorded from system UI / Start menu where UIA lookup failed)
                fx, fy = int(fallback_x), int(fallback_y)
                wininput.move_to(fx, fy)
                time.sleep(0.05)
                wininput.left_down(fx, fy)
                time.sleep(0.05)
                wininput.left_up(fx, fy)
            else:
                try:
                    _w, ctrl = _resolve_target(current_window, args)
                    uia.click_control(ctrl)
                    current_window = _w
                except Exception as uia_err:
                    if fallback_x is not None and fallback_y is not None:
                        # UIA resolve failed — fall back to recorded screen coordinates
                        # (common for WebView2/Electron apps where inner elements aren't exposed)
                        print(f"  [click] UIA resolve failed, falling back to screen coords ({fallback_x},{fallback_y})")
                        fx, fy = int(fallback_x), int(fallback_y)
                        wininput.move_to(fx, fy)
                        time.sleep(0.05)
                        wininput.left_down(fx, fy)
                        time.sleep(0.05)
                        wininput.left_up(fx, fy)
                    else:
                        raise

        elif a == "type":
            _w, ctrl = _resolve_target(current_window, args)
            uia.type_into(ctrl, text=str(args.get("text", "")), enter=bool(args.get("enter", False)))
            current_window = _w

        elif a == "sleep":
            time.sleep(float(args.get("seconds", 1)))

        # Browser steps (Playwright)
        elif a == "browser-open":
            from . import browser as browser_mod

            browser_mod.open_url(
                url=str(args["url"]),
                browser=str(args.get("browser", "chromium")),
                headless=bool(args.get("headless", False)),
            )

        elif a == "browser-navigate":
            from . import browser as browser_mod

            browser_mod.navigate(url=str(args["url"]))

        elif a == "browser-click":
            from . import browser as browser_mod

            browser_mod.click(
                selector=args.get("selector"),
                text=args.get("text"),
                exact=bool(args.get("exact", False)),
            )

        elif a == "browser-type":
            from . import browser as browser_mod

            browser_mod.type_text(
                selector=str(args["selector"]),
                text=str(args.get("text", "")),
                clear=bool(args.get("clear", True)),
            )

        elif a == "browser-eval":
            from . import browser as browser_mod

            browser_mod.evaluate(js=str(args["js"]))

        else:
            raise ValueError(f"Unknown action: {a}")


def _resolve_target(current_window, args: dict[str, Any]):
    """Resolve a target control either via classic find args or via a recorded selector."""

    timeout = int(args.get("timeout", 20))

    # Recorded selector mode: args.selector = { window: {...}, targets: [...] }
    selector = args.get("selector")
    if selector:
        win_spec = selector.get("window") or {}
        win_title_regex = args.get("window_title_regex") or win_spec.get("title_regex")
        if win_title_regex:
            w = uia.focus_window(title_regex=win_title_regex)
        else:
            win_title = win_spec.get("title")
            win_handle = win_spec.get("handle")
            if win_title:
                w = uia.focus_window(title=win_title)
            elif win_handle:
                # Passive recorder captures handle — use it as fallback
                try:
                    w = uia.focus_window(handle=int(win_handle))
                except Exception:
                    if current_window is None:
                        raise RuntimeError("Selector step needs a window. Provide window_title_regex or add a focus step.")
                    w = current_window
            else:
                if current_window is None:
                    raise RuntimeError("Selector step needs a window. Provide window_title_regex or add a focus step.")
                w = current_window

        ctrl = uia.resolve_selector(w, selector)
        return w, ctrl

    # Multi-candidate recorded selectors (preferred)
    selector_candidates = args.get("selector_candidates")
    if selector_candidates:
        if not isinstance(selector_candidates, list):
            raise ValueError("selector_candidates must be a list")

        last_err: Exception | None = None
        for sel in selector_candidates:
            try:
                if not isinstance(sel, dict):
                    continue
                w, ctrl = _resolve_target(current_window, {"selector": sel, **args})
                return w, ctrl
            except Exception as e:
                last_err = e
                continue

        raise LookupError(f"Failed to resolve selector_candidates. Last error: {last_err}")

    # Classic (manual) mode
    if current_window is None:
        raise RuntimeError("No active window. Use a 'focus' step first.")

    ctrl = uia.wait_for_control(current_window, **_control_args({**args, "timeout": timeout}))
    return current_window, ctrl


def _control_args(args: dict[str, Any]) -> dict[str, Any]:
    return {
        "control": args.get("control"),
        "auto_id": args.get("auto_id"),
        "control_type": args.get("control_type"),
        "name": args.get("name"),
        "name_regex": args.get("name_regex"),
        "timeout": int(args.get("timeout", 20)),
    }
