from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable

from pywinauto.base_wrapper import BaseWrapper

from .selectors import selector_path_from_element


@dataclass
class TreeNode:
    name: str
    control_type: str
    auto_id: str | None
    class_name: str | None
    rectangle: str
    children: list["TreeNode"]

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "control_type": self.control_type,
            "auto_id": self.auto_id,
            "class_name": self.class_name,
            "rectangle": self.rectangle,
            "children": [c.to_dict() for c in self.children],
        }


def build_tree(root: BaseWrapper, depth: int = 3, max_nodes: int = 5000) -> TreeNode:
    count = 0

    def rec(elem: BaseWrapper, d: int) -> TreeNode:
        nonlocal count
        count += 1
        info = elem.element_info
        node = TreeNode(
            name=str(info.name or ""),
            control_type=str(info.control_type or ""),
            auto_id=str(getattr(info, "automation_id", "")) or None,
            class_name=str(getattr(info, "class_name", "")) or None,
            rectangle=str(info.rectangle),
            children=[],
        )
        if d <= 0 or count >= max_nodes:
            return node

        try:
            kids = elem.children()
        except Exception:
            kids = []

        for k in kids:
            if count >= max_nodes:
                break
            node.children.append(rec(k, d - 1))

        return node

    return rec(root, depth)


def iter_elements(root: BaseWrapper, depth: int = 3, max_nodes: int = 5000) -> Iterable[BaseWrapper]:
    """Depth-limited DFS over UIA elements."""
    count = 0

    def walk(elem: BaseWrapper, d: int):
        nonlocal count
        if count >= max_nodes:
            return
        count += 1
        yield elem
        if d <= 0:
            return
        try:
            kids = elem.children()
        except Exception:
            kids = []
        for k in kids:
            yield from walk(k, d - 1)

    yield from walk(root, depth)


def element_path_dict(elem: BaseWrapper, window_root: BaseWrapper) -> list[dict[str, Any]]:
    return [s.to_dict() for s in selector_path_from_element(elem, window_root)]
