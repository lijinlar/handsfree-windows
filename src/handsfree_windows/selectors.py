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

    Prefer auto_id when available; otherwise fall back to name/control_type and sibling index.
    """

    chain: list[BaseWrapper] = []
    cur = elem
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

        idx = None
        try:
            siblings = parent.children()
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
                auto_id=str(getattr(info, "automation_id", "")) or None,
                class_name=str(getattr(info, "class_name", "")) or None,
                index=idx,
            )
        )

    return path


def candidate_targets_for_element(elem: BaseWrapper, window_root: BaseWrapper) -> list[dict[str, Any]]:
    """Return ranked candidate target selectors for an element.

    Highest stability first:
    1) auto_id + control_type
    2) name + control_type
    3) full path steps (fallback)

    No app-specific logic is used.
    """

    info = elem.element_info
    control_type = str(info.control_type) if info.control_type else None
    name = str(info.name) if info.name else None
    auto_id = str(getattr(info, "automation_id", "")) or None
    class_name = str(getattr(info, "class_name", "")) or None

    out: list[dict[str, Any]] = []
    if auto_id and control_type:
        out.append({"auto_id": auto_id, "control_type": control_type})
    if name and control_type:
        out.append({"name": name, "control_type": control_type})

    # Always include path as last resort
    path = selector_path_from_element(elem, window_root)
    out.append({"path": [s.to_dict() for s in path]})

    # Add class_name as extra hint (non-binding) when present
    for t in out:
        if class_name and "class_name" not in t:
            t["class_name"] = class_name

    return out


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
        try:
            if step.auto_id and step.control_type:
                cur = (
                    cur.child_window(auto_id=step.auto_id, control_type=step.control_type).wrapper_object()
                )
                continue
            if step.name and step.control_type:
                cur = cur.child_window(title=step.name, control_type=step.control_type).wrapper_object()
                continue
        except Exception:
            pass

        try:
            children = cur.children()
        except Exception:
            children = []

        matches = [c for c in children if _match(c, step)]
        if not matches:
            raise LookupError(f"No child matched selector step: {step}")

        if step.index is not None:
            cur = matches[min(step.index, len(matches) - 1)]
        else:
            cur = matches[0]

    return cur
