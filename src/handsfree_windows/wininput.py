from __future__ import annotations

"""Windows-native input injection using SendInput.

This is more "real" than some higher-level mouse event APIs and can work better
with apps like Paint/canvas surfaces.

All coordinates are absolute screen coordinates.
"""

import ctypes
from dataclasses import dataclass

user32 = ctypes.windll.user32


@dataclass
class Point:
    x: int
    y: int


# SendInput structures
ULONG_PTR = ctypes.c_size_t


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ULONG_PTR),
    ]


class INPUT_I(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("ii", INPUT_I)]


INPUT_MOUSE = 0

MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_ABSOLUTE = 0x8000


def _screen_size() -> tuple[int, int]:
    # Use primary screen metrics. (We can enhance to virtual screen later.)
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    return int(user32.GetSystemMetrics(SM_CXSCREEN)), int(user32.GetSystemMetrics(SM_CYSCREEN))


def _to_absolute(x: int, y: int) -> tuple[int, int]:
    """Convert pixel coords to SendInput absolute coords (0..65535)."""
    w, h = _screen_size()
    w = max(1, w)
    h = max(1, h)
    ax = int(x * 65535 / (w - 1))
    ay = int(y * 65535 / (h - 1))
    return ax, ay


def _send_mouse(flags: int, x: int, y: int) -> None:
    ax, ay = _to_absolute(x, y)
    inp = INPUT()
    inp.type = INPUT_MOUSE
    inp.ii.mi = MOUSEINPUT(
        dx=ax,
        dy=ay,
        mouseData=0,
        dwFlags=flags | MOUSEEVENTF_ABSOLUTE,
        time=0,
        dwExtraInfo=ULONG_PTR(0),
    )
    n = user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
    if n != 1:
        raise OSError(f"SendInput failed: {ctypes.get_last_error()}")


def move_to(x: int, y: int) -> None:
    _send_mouse(MOUSEEVENTF_MOVE, x, y)


def left_down(x: int, y: int) -> None:
    _send_mouse(MOUSEEVENTF_MOVE | MOUSEEVENTF_LEFTDOWN, x, y)


def left_up(x: int, y: int) -> None:
    _send_mouse(MOUSEEVENTF_MOVE | MOUSEEVENTF_LEFTUP, x, y)


def drag_left(
    start_x: int,
    start_y: int,
    end_x: int,
    end_y: int,
    duration_ms: int = 800,
    steps: int = 80,
    pre_hold_ms: int = 120,
    post_hold_ms: int = 60,
) -> None:
    import time

    steps = max(1, int(steps))
    duration_ms = max(0, int(duration_ms))
    sleep_s = (duration_ms / 1000.0) / steps if steps else 0

    move_to(start_x, start_y)
    time.sleep(0.02)
    left_down(start_x, start_y)
    time.sleep(max(0, int(pre_hold_ms)) / 1000.0)

    for i in range(1, steps + 1):
        x = int(start_x + (end_x - start_x) * (i / steps))
        y = int(start_y + (end_y - start_y) * (i / steps))
        move_to(x, y)
        if sleep_s:
            time.sleep(sleep_s)

    time.sleep(max(0, int(post_hold_ms)) / 1000.0)
    left_up(end_x, end_y)
