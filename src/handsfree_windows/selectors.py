from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable

from pywinauto.base_wrapper import BaseWrapper


@dataclass(frozen=True)
class SelectorStep:
    """A single hop in a control path."""

    control_type: str | None = None
    name: str | None = None
    auto_id: str | None = None
    class_name: str | None = None
    index: int | None = None

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.control_type:
            d["control_type"] = self.control_type
        if self.name:
            d["name"] = self.name
        if self.auto_id:
            d["auto_id"] = self.auto_id
        if self.class_name:
            d["class_name"] = self.class_name
        if self.index is not None:
            d["index"] = int(self.index)
        return d

    @staticmethod
    def from_dict(d: dict[str, Any]) -> "SelectorStep":
        return SelectorStep(
            control_type=d.get("control_type"),
            name=d.get("name"),
            auto_id=d.get("auto_id"),
            class_name=d.get("class_name"),
            index=d.get("index"),
        )


def selector_path_from_element(elem: BaseWrapper, window_root: BaseWrapper) -> list[SelectorStep]:
    """Return a best-effort stable path from window_root -> elem.

    The path is meant to be robust-ish across sessions, but it is not guaranteed.
    Prefer auto_id when available; otherwise fall back to name/control_type and sibling index.
    """

    # Walk up via element_info.parent (UIA tree), then back down.
    chain: list[BaseWrapper] = []
    cur = elem
    # Guard against weird trees
    for _ in range(256):
        chain.append(cur)
        if int(cur.handle) == int(window_root.handle):
            break
        parent = getattr(cur.element_info, "parent", None)
        if parent is None:
            break
        try:
            cur = parent.wrapper_object()
        except Exception:
            break

    if not chain or int(chain[-1].handle) != int(window_root.handle):
        raise ValueError("Element is not within the given window root")

    chain = list(reversed(chain))  # root -> elem

    path: list[SelectorStep] = []
    for i in range(1, len(chain)):
        node = chain[i]
        info = node.element_info
        parent = chain[i - 1]

        # Determine sibling index among similar-ish children to disambiguate
        idx = None
        try:
            siblings = parent.children()
            # Try exact handle match in siblings
            for j, sib in enumerate(siblings):
                if int(sib.handle) == int(node.handle):
                    idx = j
                    break
        except Exception:
            idx = None

        path.append(
            SelectorStep(
                control_type=str(info.control_type) if info.control_type else None,
                name=str(info.name) if info.name else None,
                auto_id=str(info.automation_id) if getattr(info, "automation_id", None) else None,
                class_name=str(info.class_name) if getattr(info, "class_name", None) else None,
                index=idx,
            )
        )

    return path


def _match(child: BaseWrapper, step: SelectorStep) -> bool:
    info = child.element_info
    if step.control_type and str(info.control_type) != step.control_type:
        return False
    if step.auto_id and str(getattr(info, "automation_id", "")) != step.auto_id:
        return False
    if step.class_name and str(getattr(info, "class_name", "")) != step.class_name:
        return False
    if step.name and str(info.name or "") != step.name:
        return False
    return True


def resolve_selector_path(window_root: BaseWrapper, path: Iterable[SelectorStep]) -> BaseWrapper:
    cur = window_root
    for step in path:
        # First try direct pywinauto selector when possible (fast)
        try:
            if step.auto_id and step.control_type:
                cur = cur.child_window(auto_id=step.auto_id, control_type=step.control_type).wrapper_object()
                continue
            if step.name and step.control_type:
                cur = cur.child_window(title=step.name, control_type=step.control_type).wrapper_object()
                continue
        except Exception:
            # Fall back to manual scan
            pass

        children = []
        try:
            children = cur.children()
        except Exception:
            children = []

        matches = [c for c in children if _match(c, step)]
        if not matches:
            raise LookupError(f"No child matched selector step: {step}")

        if step.index is not None:
            try:
                cur = matches[0 if len(matches) == 1 else min(step.index, len(matches) - 1)]
            except Exception:
                cur = matches[0]
        else:
            cur = matches[0]

    return cur
