from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table

from . import discover, macro, tree, uia

app = typer.Typer(add_completion=False, help="Handsfree Windows: control apps via UI Automation (UIA)")
console = Console()


def _window_kwargs(
    title: Optional[str] = None,
    title_regex: Optional[str] = None,
    handle: Optional[int] = None,
):
    return {"title": title, "title_regex": title_regex, "handle": handle}


@app.command("list-windows")
def list_windows(
    title_regex: Optional[str] = typer.Option(None, help="Regex filter for window titles"),
    json_out: bool = typer.Option(False, "--json", help="Output as JSON"),
):
    """List top-level windows."""
    wins = uia.list_top_windows(title_regex=title_regex)
    if json_out:
        console.print(json.dumps([w.__dict__ for w in wins], ensure_ascii=False, indent=2))
        raise typer.Exit(0)

    t = Table(title="Top-level windows")
    t.add_column("Handle", justify="right")
    t.add_column("PID", justify="right")
    t.add_column("Class")
    t.add_column("Title")
    for w in wins:
        t.add_row(str(w.handle), str(w.pid), w.class_name, w.title)
    console.print(t)


@app.command()
def focus(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
):
    """Focus a window."""
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    console.print(f"Focused: {w.window_text()!r}")


@app.command("tree")
def export_tree(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
    depth: int = typer.Option(3, help="Tree depth to traverse"),
    max_nodes: int = typer.Option(5000, help="Max nodes to export"),
):
    """Export the UI Automation tree for a window (JSON)."""
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    tnode = tree.build_tree(w, depth=depth, max_nodes=max_nodes)
    console.print(json.dumps(tnode.to_dict(), ensure_ascii=False, indent=2))


@app.command("list-controls")
def list_controls(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
    depth: int = typer.Option(3, help="Tree depth to traverse"),
    limit: int = typer.Option(200, help="Max controls to print"),
):
    """List controls under a window (UIA tree)."""
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    t = Table(title=f"Controls (depth={depth})")
    t.add_column("Name")
    t.add_column("Type")
    t.add_column("AutoId")
    t.add_column("Class")
    t.add_column("Rect")

    count = 0
    for c in uia.iter_controls(w, depth=depth):
        t.add_row(c.name, c.control_type, c.auto_id or "", c.class_name or "", c.rectangle)
        count += 1
        if count >= limit:
            break

    console.print(t)
    console.print(f"Shown: {count}")


@app.command()
def click(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
    control: Optional[str] = typer.Option(None, help="best_match (e.g. 'OK', 'Edit')"),
    auto_id: Optional[str] = typer.Option(None, help="UIA automation id"),
    control_type: Optional[str] = typer.Option(None, help="UIA control type (e.g. 'Button')"),
    name: Optional[str] = typer.Option(None, help="UIA name/title"),
    name_regex: Optional[str] = typer.Option(None, help="Regex for UIA name/title"),
    timeout: int = typer.Option(20, help="Seconds to wait for control"),
):
    """Click a control inside a window."""
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    ctrl = uia.wait_for_control(
        w,
        timeout=timeout,
        control=control,
        auto_id=auto_id,
        control_type=control_type,
        name=name,
        name_regex=name_regex,
    )
    uia.click_control(ctrl)
    console.print("Clicked")


@app.command()
def type(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
    text: str = typer.Option(..., help="Text to type/set"),
    enter: bool = typer.Option(False, help="Press Enter after typing"),
    control: Optional[str] = typer.Option(None, help="best_match (e.g. 'Edit')"),
    auto_id: Optional[str] = typer.Option(None, help="UIA automation id"),
    control_type: Optional[str] = typer.Option(None, help="UIA control type"),
    name: Optional[str] = typer.Option(None, help="UIA name/title"),
    name_regex: Optional[str] = typer.Option(None, help="Regex for UIA name/title"),
    timeout: int = typer.Option(20, help="Seconds to wait for control"),
):
    """Type into a control."""
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    ctrl = uia.wait_for_control(
        w,
        timeout=timeout,
        control=control,
        auto_id=auto_id,
        control_type=control_type,
        name=name,
        name_regex=name_regex,
    )
    uia.type_into(ctrl, text=text, enter=enter)
    console.print("Typed")


@app.command("start")
def start_menu_launch(
    app_name: str = typer.Option(..., "--app", help="Application name as shown in Start menu search"),
    delay_ms: int = typer.Option(250, help="Delay after opening Start (ms)"),
):
    """Open Start menu, search app_name, and launch it.

    This is intentionally simple (human-style): Win key -> type -> Enter.
    """
    from pywinauto.keyboard import send_keys

    send_keys("{VK_LWIN}")
    import time

    time.sleep(max(0, delay_ms) / 1000.0)
    send_keys(app_name, with_spaces=True)
    time.sleep(0.1)
    send_keys("{ENTER}")
    console.print(f"Launched (start menu): {app_name}")


@app.command("open-path")
def open_path(
    path: str = typer.Option(..., help="Filesystem path to open in Explorer"),
    direct: bool = typer.Option(True, help="Use explorer.exe <path> (most reliable)"),
    use_win_e: bool = typer.Option(True, help="If not direct: open Explorer via Win+E first"),
    delay_ms: int = typer.Option(600, help="If not direct: delay after opening Explorer (ms)"),
):
    """Open File Explorer and navigate to a folder.

    Two generic strategies:
    - direct (default): run `explorer.exe <path>`
    - human-style: Win+E -> Ctrl+L -> paste path -> Enter
    """

    if direct:
        import subprocess

        # explorer.exe is the canonical way to open a folder reliably.
        subprocess.Popen(["explorer.exe", path])
        console.print(f"Opened path in Explorer (direct): {path}")
        return

    from pywinauto.keyboard import send_keys
    import time

    if use_win_e:
        send_keys("{VK_LWIN down}e{VK_LWIN up}")
        time.sleep(max(0, delay_ms) / 1000.0)

    # Ensure we're not stuck in the breadcrumbs/search UI
    send_keys("{ESC}")
    time.sleep(0.05)

    # Focus address bar
    send_keys("^l")
    time.sleep(0.1)

    # Paste path via clipboard
    try:
        import pyperclip  # type: ignore

        pyperclip.copy(path)
    except Exception:
        import ctypes

        CF_UNICODETEXT = 13
        GMEM_MOVEABLE = 0x0002

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        if not user32.OpenClipboard(None):
            raise RuntimeError("Unable to open clipboard")
        try:
            user32.EmptyClipboard()
            data = (path + "\0").encode("utf-16le")
            h_global = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
            if not h_global:
                raise RuntimeError("GlobalAlloc failed")
            lp = kernel32.GlobalLock(h_global)
            if not lp:
                raise RuntimeError("GlobalLock failed")
            try:
                ctypes.memmove(lp, data, len(data))
            finally:
                kernel32.GlobalUnlock(h_global)
            user32.SetClipboardData(CF_UNICODETEXT, h_global)
        finally:
            user32.CloseClipboard()

    send_keys("^v")
    time.sleep(0.1)
    send_keys("{ENTER}")

    console.print(f"Opened path in Explorer (human): {path}")


@app.command("inspect")
def inspect_under_cursor(
    json_out: bool = typer.Option(False, "--json", help="Output selector as JSON"),
):
    """Inspect the UI element under the mouse cursor and print a robust selector."""
    x, y = uia.cursor_pos()
    elem = uia.element_from_point(x, y)
    sel = uia.selector_for_element(elem)
    if json_out:
        console.print(json.dumps(sel, ensure_ascii=False, indent=2))
    else:
        console.print(f"Cursor: ({x}, {y})")
        console.print(json.dumps(sel, ensure_ascii=False, indent=2))


@app.command("resolve")
def resolve_selector_cmd(
    selector_json: Optional[str] = typer.Option(None, help="Selector JSON string"),
    selector_file: Optional[Path] = typer.Option(None, exists=True, help="Path to selector JSON"),
    title_regex: Optional[str] = typer.Option(None, help="Override window title regex"),
):
    """Resolve a selector against the current UI and print what it matched."""
    if not selector_json and not selector_file:
        raise typer.BadParameter("Provide --selector-json or --selector-file")

    if selector_file:
        data = json.loads(selector_file.read_text(encoding="utf-8"))
    else:
        data = json.loads(selector_json or "{}")

    if title_regex:
        data.setdefault("window", {})["title_regex"] = title_regex

    wspec = data.get("window") or {}
    if wspec.get("title_regex"):
        w = uia.focus_window(title_regex=wspec["title_regex"])
    elif wspec.get("title"):
        w = uia.focus_window(title=wspec["title"])
    else:
        raise typer.BadParameter("Selector must include window.title or window.title_regex")

    ctrl = uia.resolve_selector(w, data)
    info = ctrl.element_info
    out = {
        "matched": {
            "name": str(info.name or ""),
            "control_type": str(info.control_type or ""),
            "auto_id": str(getattr(info, "automation_id", "") or ""),
            "class_name": str(getattr(info, "class_name", "") or ""),
            "rectangle": str(info.rectangle),
        }
    }
    console.print(json.dumps(out, ensure_ascii=False, indent=2))


@app.command("record")
def record_macro(
    out: Path = typer.Option(Path("macro.yaml"), help="Output macro YAML"),
    window_title_regex: Optional[str] = typer.Option(
        None, help="If set, force focus to this window before each step"
    ),
):
    """Interactive macro recorder (generic).

    Workflow:
    - Hover over a UI element
    - Choose an action (click/type)
    - The CLI records multiple selector candidates for that element.

    Stop by entering 'q'.
    """

    steps: list[dict] = []

    console.print(
        "Recording. For each step: hover a UI element, then choose action. Enter 'q' to finish."
    )

    while True:
        action = console.input("Action [click/type/sleep/q]: ").strip().lower()
        if action in {"q", "quit", "exit"}:
            break

        if action == "sleep":
            seconds = float(console.input("Seconds: ").strip() or "1")
            steps.append({"action": "sleep", "args": {"seconds": seconds}})
            continue

        if action not in {"click", "type"}:
            console.print("Unknown action. Use click/type/sleep/q.")
            continue

        x, y = uia.cursor_pos()
        elem = uia.element_from_point(x, y)
        sel = uia.selector_for_element(elem)

        # Store the same selector multiple times is pointless; but we allow expansion later.
        selector_candidates = [sel]

        args = {
            "selector_candidates": selector_candidates,
            "timeout": 20,
        }

        if window_title_regex:
            # Encourage regex matching for skills
            for s in selector_candidates:
                s.setdefault("window", {})["title_regex"] = window_title_regex
            args["window_title_regex"] = window_title_regex

        if action == "type":
            text = console.input("Text: ")
            enter = console.input("Press Enter after? [y/N]: ").strip().lower() in {"y", "yes"}
            args["text"] = text
            args["enter"] = enter

        steps.append({"action": action, "args": args})
        console.print(f"Recorded {action} at cursor ({x},{y})")

    import yaml

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(yaml.safe_dump(steps, sort_keys=False, allow_unicode=True), encoding="utf-8")
    console.print(f"Saved macro: {out}")


@app.command("click-at")
def click_at_cmd(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
    x: int = typer.Option(..., help="X offset (window-relative)"),
    y: int = typer.Option(..., help="Y offset (window-relative)"),
):
    """Click at window-relative coordinates."""
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    uia.click_at(w, x=x, y=y)
    console.print(f"Clicked at ({x},{y})")


@app.command("canvas-selector")
def canvas_selector(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
):
    """Heuristic: output selector JSON for the largest pane/custom/document inside a window."""
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    sel = discover.selector_for_largest_pane(w)
    console.print(json.dumps(sel, ensure_ascii=False, indent=2))


@app.command("drag")
def drag_cmd(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
    start_x: int = typer.Option(..., help="Start X (window-relative)"),
    start_y: int = typer.Option(..., help="Start Y (window-relative)"),
    end_x: int = typer.Option(..., help="End X (window-relative)"),
    end_y: int = typer.Option(..., help="End Y (window-relative)"),
    duration_ms: int = typer.Option(400, help="Drag duration in ms"),
    steps: int = typer.Option(30, help="Interpolation steps"),
):
    """Drag mouse from start->end using window-relative coords."""
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    uia.drag(
        w,
        start_x=start_x,
        start_y=start_y,
        end_x=end_x,
        end_y=end_y,
        duration_ms=duration_ms,
        steps=steps,
    )
    console.print(f"Dragged ({start_x},{start_y}) -> ({end_x},{end_y})")


@app.command("drag-screen")
def drag_screen_cmd(
    start_x: int = typer.Option(..., help="Start X (screen)"),
    start_y: int = typer.Option(..., help="Start Y (screen)"),
    end_x: int = typer.Option(..., help="End X (screen)"),
    end_y: int = typer.Option(..., help="End Y (screen)"),
    duration_ms: int = typer.Option(600, help="Drag duration in ms"),
    steps: int = typer.Option(40, help="Interpolation steps"),
):
    """Drag using absolute screen coordinates."""
    uia.drag_screen(
        start_x=start_x,
        start_y=start_y,
        end_x=end_x,
        end_y=end_y,
        duration_ms=duration_ms,
        steps=steps,
    )
    console.print(f"Dragged screen ({start_x},{start_y}) -> ({end_x},{end_y})")


@app.command("run")
def run_macro_cmd(path: Path = typer.Argument(..., exists=True)):
    """Run a YAML macro."""
    macro.run_macro(path)
    console.print(f"Done: {path}")


if __name__ == "__main__":
    app()
