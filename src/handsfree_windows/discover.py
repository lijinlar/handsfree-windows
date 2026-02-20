from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional

from pywinauto.base_wrapper import BaseWrapper

from . import tree, uia


@dataclass
class Rect:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return max(0, self.right - self.left)

    @property
    def height(self) -> int:
        return max(0, self.bottom - self.top)

    def inset(self, pad: int) -> "Rect":
        return Rect(
            left=self.left + pad,
            top=self.top + pad,
            right=self.right - pad,
            bottom=self.bottom - pad,
        )


def _rect_from_str(s: str) -> Optional[Rect]:
    import re

    nums = [int(x) for x in re.findall(r"-?\d+", s or "")]
    if len(nums) >= 4:
        return Rect(nums[0], nums[1], nums[2], nums[3])
    return None


def largest_child_pane(window: BaseWrapper, depth: int = 12, max_nodes: int = 40000) -> tuple[BaseWrapper, Rect]:
    """Find the largest descendant element of type Pane/Custom/Document.

    This is a generic heuristic that often corresponds to the main content/canvas area.
    """

    best_elem: BaseWrapper | None = None
    best_rect: Rect | None = None
    best_area = -1

    for elem in tree.iter_elements(window, depth=depth, max_nodes=max_nodes):
        try:
            info = elem.element_info
            ct = str(info.control_type or "")
            if ct not in {"Pane", "Custom", "Document"}:
                continue
            r = _rect_from_str(str(info.rectangle))
            if not r:
                continue
            area = r.width * r.height
            if area > best_area:
                best_area = area
                best_elem = elem
                best_rect = r
        except Exception:
            continue

    if not best_elem or not best_rect:
        raise LookupError("No suitable pane/custom/document element found")

    return best_elem, best_rect


def selector_for_largest_pane(window: BaseWrapper) -> dict[str, Any]:
    elem, _ = largest_child_pane(window)
    return uia.selector_for_element(elem)
