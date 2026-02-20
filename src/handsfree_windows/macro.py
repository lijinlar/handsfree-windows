from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from . import uia


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

    for step in steps:
        a = step.action
        args = step.args

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
            import time

            time.sleep(max(0, delay_ms) / 1000.0)
            send_keys(app_name, with_spaces=True)
            time.sleep(0.1)
            send_keys("{ENTER}")

        elif a == "click":
            _w, ctrl = _resolve_target(current_window, args)
            uia.click_control(ctrl)
            current_window = _w

        elif a == "type":
            _w, ctrl = _resolve_target(current_window, args)
            uia.type_into(ctrl, text=str(args.get("text", "")), enter=bool(args.get("enter", False)))
            current_window = _w

        elif a == "sleep":
            import time

            time.sleep(float(args.get("seconds", 1)))

        else:
            raise ValueError(f"Unknown action: {a}")


def _resolve_target(current_window, args: dict[str, Any]):
    """Resolve a target control either via classic find args or via a recorded selector."""

    timeout = int(args.get("timeout", 20))

    # Recorded selector mode: args.selector = { window: {...}, path: [...] }
    selector = args.get("selector")
    if selector:
        win_title_regex = args.get("window_title_regex")
        if win_title_regex:
            w = uia.focus_window(title_regex=win_title_regex)
        else:
            # Best-effort: focus by exact title if present
            win_title = (selector.get("window") or {}).get("title")
            if win_title:
                w = uia.focus_window(title=win_title)
            else:
                if current_window is None:
                    raise RuntimeError("Selector step needs a window. Provide window_title_regex or add a focus step.")
                w = current_window

        # Resolve the path
        path = selector.get("path")
        if not isinstance(path, list):
            raise ValueError("selector.path must be a list")

        ctrl = uia.resolve_selector(w, path)
        return w, ctrl

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
