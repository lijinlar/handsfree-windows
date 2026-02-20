from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable, Optional

from pywinauto import Desktop
from pywinauto.base_wrapper import BaseWrapper
from pywinauto.findwindows import ElementNotFoundError


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


def click_control(ctrl: BaseWrapper) -> None:
    ctrl.wait("enabled", timeout=10)
    try:
        ctrl.click_input()
    except Exception:
        # Fallback for some controls
        ctrl.invoke()


def type_into(ctrl: BaseWrapper, text: str, enter: bool = False) -> None:
    ctrl.wait("enabled", timeout=10)
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
