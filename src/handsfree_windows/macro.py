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

        elif a == "click":
            if current_window is None:
                raise RuntimeError("No active window. Use a 'focus' step first.")
            ctrl = uia.wait_for_control(current_window, **_control_args(args))
            uia.click_control(ctrl)

        elif a == "type":
            if current_window is None:
                raise RuntimeError("No active window. Use a 'focus' step first.")
            ctrl = uia.wait_for_control(current_window, **_control_args(args))
            uia.type_into(ctrl, text=str(args.get("text", "")), enter=bool(args.get("enter", False)))

        elif a == "sleep":
            import time

            time.sleep(float(args.get("seconds", 1)))

        else:
            raise ValueError(f"Unknown action: {a}")


def _control_args(args: dict[str, Any]) -> dict[str, Any]:
    return {
        "control": args.get("control"),
        "auto_id": args.get("auto_id"),
        "control_type": args.get("control_type"),
        "name": args.get("name"),
        "name_regex": args.get("name_regex"),
        "timeout": int(args.get("timeout", 20)),
    }
