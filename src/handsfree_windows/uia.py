from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable, Optional

from pywinauto import Desktop
from pywinauto.base_wrapper import BaseWrapper
from pywinauto.findwindows import ElementNotFoundError

from .selectors import (
    SelectorStep,
    candidate_targets_for_element,
    resolve_selector_path,
    selector_path_from_element,
)


@dataclass
class WindowSpec:
    handle: int
    title: str
    class_name: str
    pid: int


def _desktop() -> Desktop:
    # UIA backend is the most broadly compatible for modern Windows apps.
    return Desktop(backend="uia")


def list_top_windows(title_regex: str | None = None) -> list[WindowSpec]:
    pattern = re.compile(title_regex) if title_regex else None
    out: list[WindowSpec] = []
    for w in _desktop().windows():
        try:
            title = w.window_text() or ""
            if pattern and not pattern.search(title):
                continue
            out.append(
                WindowSpec(
                    handle=int(w.handle),
                    title=title,
                    class_name=w.friendly_class_name(),
                    pid=int(w.process_id()),
                )
            )
        except Exception:
            # Some windows can be transient/permission-limited.
            continue
    return out


def get_window(title: str | None = None, title_regex: str | None = None, handle: int | None = None):
    d = _desktop()
    if handle is not None:
        return d.window(handle=handle)

    if title is not None:
        return d.window(title=title)

    if title_regex is not None:
        # pywinauto supports regex via title_re
        return d.window(title_re=title_regex)

    raise ValueError("Provide one of: title, title_regex, handle")


def focus_window(**kwargs) -> BaseWrapper:
    w = get_window(**kwargs)
    w.set_focus()
    return w


@dataclass
class ControlSpec:
    name: str
    control_type: str
    auto_id: str | None
    class_name: str | None
    rectangle: str


def iter_controls(window: BaseWrapper, depth: int = 3) -> Iterable[ControlSpec]:
    # UIA tree walk
    def walk(elem: BaseWrapper, d: int) -> Iterable[ControlSpec]:
        try:
            info = elem.element_info
            yield ControlSpec(
                name=str(info.name or ""),
                control_type=str(info.control_type or ""),
                auto_id=str(info.automation_id) if info.automation_id else None,
                class_name=str(info.class_name) if info.class_name else None,
                rectangle=str(info.rectangle),
            )
            if d <= 0:
                return
            for child in elem.children():
                yield from walk(child, d - 1)
        except Exception:
            return

    yield from walk(window, depth)


def find_control(
    window: BaseWrapper,
    control: str | None = None,
    auto_id: str | None = None,
    control_type: str | None = None,
    name: str | None = None,
    name_regex: str | None = None,
) -> BaseWrapper:
    """Find a control within a window.

    Matching options:
    - control: pywinauto best_match (e.g. "OK", "Edit", "Save")
    - auto_id + control_type: more deterministic when available
    - name/name_regex + control_type
    """

    if control:
        return window.child_window(best_match=control)

    if auto_id:
        return window.child_window(auto_id=auto_id, control_type=control_type)

    if name_regex:
        return window.child_window(title_re=name_regex, control_type=control_type)

    if name:
        return window.child_window(title=name, control_type=control_type)

    raise ValueError("Provide one of: control, auto_id, name, name_regex")


def _wait_enabled(ctrl: BaseWrapper, timeout: int = 10) -> None:
    import time

    end = time.time() + timeout
    while time.time() < end:
        try:
            if getattr(ctrl, "is_enabled", None) and ctrl.is_enabled():
                return
        except Exception:
            pass
        time.sleep(0.1)


def click_control(ctrl: BaseWrapper) -> None:
    _wait_enabled(ctrl, timeout=10)
    try:
        ctrl.click_input()
    except Exception:
        # Fallback for some controls
        try:
            ctrl.invoke()
        except Exception:
            # Last resort: click at center
            r = ctrl.rectangle()
            from pywinauto import mouse

            mouse.click(coords=(int(r.mid_point().x), int(r.mid_point().y)))


def type_into(ctrl: BaseWrapper, text: str, enter: bool = False) -> None:
    _wait_enabled(ctrl, timeout=10)
    try:
        ctrl.set_focus()
    except Exception:
        pass

    # Some controls support set_edit_text, others require type_keys.
    try:
        ctrl.set_edit_text(text)
    except Exception:
        from pywinauto.keyboard import send_keys

        send_keys(text, with_spaces=True)

    if enter:
        from pywinauto.keyboard import send_keys

        send_keys("{ENTER}")


def wait_for_control(window: BaseWrapper, timeout: int = 20, **find_kwargs) -> BaseWrapper:
    import time

    end = time.time() + timeout
    last_err: Optional[Exception] = None
    while time.time() < end:
        try:
            ctrl = find_control(window, **find_kwargs)
            # Resolve wrapper
            ctrl = ctrl.wrapper_object()
            return ctrl
        except (ElementNotFoundError, Exception) as e:
            last_err = e
            time.sleep(0.5)
    raise TimeoutError(f"Control not found within {timeout}s. Last error: {last_err}")


def element_from_point(x: int, y: int) -> BaseWrapper:
    """Get the UIA element under screen coordinates (x, y)."""
    return _desktop().from_point(x, y)


def cursor_pos() -> tuple[int, int]:
    """Current cursor position (screen coords). Uses win32api if available."""
    try:
        import win32api  # type: ignore

        x, y = win32api.GetCursorPos()
        return int(x), int(y)
    except Exception:
        # Fallback via ctypes
        import ctypes

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        pt = POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        return int(pt.x), int(pt.y)


def top_level_window_for(elem: BaseWrapper) -> BaseWrapper:
    """Find the nearest top-level window ancestor for an element."""
    cur = elem
    for _ in range(256):
        parent = getattr(cur.element_info, "parent", None)
        if parent is None:
            break
        try:
            pwrap = parent.wrapper_object()
        except Exception:
            break

        # A top-level window typically has no parent, or its parent is the Desktop.
        # pywinauto wrapper has .top_level_parent(), but it can fail for some elements.
        try:
            tl = cur.top_level_parent()
            if tl is not None:
                return tl
        except Exception:
            pass

        cur = pwrap

    # Last resort: return element itself
    return elem


def selector_for_element(elem: BaseWrapper) -> dict:
    """Build a selector payload for an element.

    Returns an app-agnostic selector with multiple candidate targeting strategies.

    Schema (v2):
    {
      "window": {"title": "...", "title_regex": "..."},
      "targets": [
        {"auto_id": "...", "control_type": "Button"},
        {"name": "OK", "control_type": "Button"},
        {"path": [ ... SelectorStep dicts ... ]}
      ]
    }

    We include window.title as observed; for skills, prefer using window.title_regex.
    """

    win = top_level_window_for(elem)
    try:
        win_title = win.window_text()
    except Exception:
        win_title = ""

    targets = candidate_targets_for_element(elem, win)

    handle = win.handle
    try:
        pid = int(getattr(win, "process_id", lambda: None)())
    except (TypeError, ValueError):
        pid = None
    return {
        "window": {
            "title": win_title,
            "handle": int(handle) if handle is not None else None,
            "pid": pid,
        },
        "targets": targets,
    }


def resolve_selector(window: BaseWrapper, selector: dict) -> BaseWrapper:
    """Resolve a v2 selector against a window.

    Tries selector["targets"] in order. Each target can be:
    - {auto_id, control_type}
    - {name, control_type}
    - {path: [...steps...]}
    """

    targets = selector.get("targets") or []
    if not isinstance(targets, list) or not targets:
        raise ValueError("selector.targets must be a non-empty list")

    last_err: Exception | None = None

    for t in targets:
        try:
            if not isinstance(t, dict):
                continue

            if t.get("auto_id") and t.get("control_type"):
                return window.child_window(
                    auto_id=t["auto_id"], control_type=t["control_type"]
                ).wrapper_object()

            if t.get("name") and t.get("control_type"):
                return window.child_window(title=t["name"], control_type=t["control_type"]).wrapper_object()

            if t.get("path"):
                steps = [SelectorStep.from_dict(p) for p in t["path"]]
                return resolve_selector_path(window, steps)

        except Exception as e:
            last_err = e
            continue

    raise LookupError(f"Failed to resolve selector targets. Last error: {last_err}")


def window_rect(window: BaseWrapper) -> tuple[int, int, int, int]:
    """Return window rectangle (left, top, right, bottom) in screen coords."""
    r = window.rectangle()
    return int(r.left), int(r.top), int(r.right), int(r.bottom)


def client_point(window: BaseWrapper, x: int, y: int) -> tuple[int, int]:
    """Convert window-client relative coords to screen coords.

    This uses the top-left of the window rectangle as an approximation. For many apps
    this is good enough for human-like automation.
    """
    l, t, _, _ = window_rect(window)
    return l + int(x), t + int(y)


def click_at(window: BaseWrapper, x: int, y: int, button: str = "left") -> None:
    """Click at window-relative coordinates."""
    from pywinauto import mouse

    sx, sy = client_point(window, x, y)
    mouse.click(button=button, coords=(sx, sy))


def click_screen(x: int, y: int, button: str = "left") -> None:
    """Click at absolute screen coordinates."""
    from pywinauto import mouse

    mouse.click(button=button, coords=(int(x), int(y)))


def drag_screen(
    start_x: int,
    start_y: int,
    end_x: int,
    end_y: int,
    button: str = "left",
    duration_ms: int = 300,
    steps: int = 25,
    pre_hold_ms: int = 120,
    post_hold_ms: int = 50,
    backend: str = "pywinauto",
) -> None:
    """Drag using absolute screen coordinates.

    Some apps (notably Paint / ink surfaces) need a small dwell time after mouse down
    before movement is recognized as a drag.
    """
    sx, sy = int(start_x), int(start_y)
    ex, ey = int(end_x), int(end_y)

    backend = (backend or "pywinauto").lower()

    if backend == "sendinput":
        from . import wininput

        if button != "left":
            raise ValueError("sendinput backend currently supports left button only")

        wininput.drag_left(
            start_x=sx,
            start_y=sy,
            end_x=ex,
            end_y=ey,
            duration_ms=duration_ms,
            steps=steps,
            pre_hold_ms=pre_hold_ms,
            post_hold_ms=post_hold_ms,
        )
        return

    # Default: pywinauto mouse helpers
    from pywinauto import mouse
    import time

    # Try builtin drag first
    if hasattr(mouse, "drag"):
        try:
            mouse.drag(coords=(sx, sy), coords2=(ex, ey))
            return
        except Exception:
            pass

    steps = max(1, int(steps))
    duration_ms = max(0, int(duration_ms))
    sleep_s = (duration_ms / 1000.0) / steps if steps else 0

    # Ensure we actually move there before pressing
    try:
        mouse.move(coords=(sx, sy))
        time.sleep(0.02)
    except Exception:
        pass

    mouse.press(button=button, coords=(sx, sy))
    time.sleep(max(0, int(pre_hold_ms)) / 1000.0)

    try:
        for i in range(1, steps + 1):
            x = int(sx + (ex - sx) * (i / steps))
            y = int(sy + (ey - sy) * (i / steps))
            mouse.move(coords=(x, y))
            if sleep_s:
                time.sleep(sleep_s)
    finally:
        time.sleep(max(0, int(post_hold_ms)) / 1000.0)
        mouse.release(button=button, coords=(ex, ey))


def drag(
    window: BaseWrapper,
    start_x: int,
    start_y: int,
    end_x: int,
    end_y: int,
    button: str = "left",
    duration_ms: int = 300,
    steps: int = 25,
) -> None:
    """Drag from start->end using window-relative coordinates.

    Uses a stepped drag with small delays (more human-like; improves reliability in apps like Paint).
    """
    from pywinauto import mouse
    import time

    sx, sy = client_point(window, start_x, start_y)
    ex, ey = client_point(window, end_x, end_y)

    # Prefer built-in drag if available
    if hasattr(mouse, "drag"):
        try:
            mouse.drag(coords=(sx, sy), coords2=(ex, ey))
            return
        except Exception:
            pass

    steps = max(1, int(steps))
    duration_ms = max(0, int(duration_ms))
    sleep_s = (duration_ms / 1000.0) / steps if steps else 0

    mouse.press(button=button, coords=(sx, sy))
    try:
        for i in range(1, steps + 1):
            x = int(sx + (ex - sx) * (i / steps))
            y = int(sy + (ey - sy) * (i / steps))
            mouse.move(coords=(x, y))
            if sleep_s:
                time.sleep(sleep_s)
    finally:
        mouse.release(button=button, coords=(ex, ey))
