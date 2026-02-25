from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table

from . import browser as browser_mod
from . import discover, macro, recorder, tree, uia

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
        None, help="If set, force focus to this window before each step (interactive mode only)"
    ),
    passive: bool = typer.Option(
        False, "--passive", help="Passive mode: record clicks and keystrokes automatically. Press F9 to stop."
    ),
    verbose: bool = typer.Option(
        False, "--verbose", "-v", help="Print each recorded step to stdout (passive mode only)"
    ),
):
    """Macro recorder -- interactive (default) or passive (--passive).

    Interactive mode:
    - Hover over a UI element
    - Choose an action (click/type)
    - The CLI records multiple selector candidates for that element.
    - Stop by entering 'q'.

    Passive mode (--passive):
    - Record starts immediately -- use the app normally.
    - Left-click: captures UIA element as a click step.
    - Typing: buffered and saved as type steps.
    - Enter key: flushes typing with enter=true.
    - F9: stops recording and saves the macro.
    """
    if passive:
        recorder.passive_record(out=out, verbose=verbose)
        return

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
    pre_hold_ms: int = typer.Option(140, help="Hold after mouse-down before moving (ms)"),
    post_hold_ms: int = typer.Option(60, help="Hold before mouse-up (ms)"),
    backend: str = typer.Option("pywinauto", help="Drag backend: pywinauto|sendinput"),
):
    """Drag using absolute screen coordinates."""
    uia.drag_screen(
        start_x=start_x,
        start_y=start_y,
        end_x=end_x,
        end_y=end_y,
        duration_ms=duration_ms,
        steps=steps,
        pre_hold_ms=pre_hold_ms,
        post_hold_ms=post_hold_ms,
        backend=backend,
    )
    console.print(f"Dragged screen ({start_x},{start_y}) -> ({end_x},{end_y})")


@app.command("drag-canvas")
def drag_canvas_cmd(
    title: Optional[str] = typer.Option(None, help="Exact window title"),
    title_regex: Optional[str] = typer.Option(None, help="Regex window title"),
    handle: Optional[int] = typer.Option(None, help="Window handle"),
    pad: int = typer.Option(30, help="Padding inside the detected canvas rect"),
    x1: float = typer.Option(0.15, help="Start X as fraction of canvas width (0-1)"),
    y1: float = typer.Option(0.20, help="Start Y as fraction of canvas height (0-1)"),
    x2: float = typer.Option(0.65, help="End X as fraction of canvas width (0-1)"),
    y2: float = typer.Option(0.55, help="End Y as fraction of canvas height (0-1)"),
    duration_ms: int = typer.Option(1200, help="Drag duration in ms"),
    steps: int = typer.Option(120, help="Interpolation steps"),
    pre_hold_ms: int = typer.Option(220, help="Hold after mouse-down before moving (ms)"),
    post_hold_ms: int = typer.Option(120, help="Hold before mouse-up (ms)"),
    backend: str = typer.Option("sendinput", help="Drag backend: pywinauto|sendinput"),
):
    """Drag within the largest detected canvas/content area of a window.

    This avoids guessing screen coordinates and ensures the drag starts inside the canvas.
    """
    w = uia.focus_window(**_window_kwargs(title, title_regex, handle))
    _elem, r = discover.largest_child_pane(w)
    r = r.inset(int(pad))

    def clamp01(v: float) -> float:
        return max(0.0, min(1.0, float(v)))

    sx = int(r.left + clamp01(x1) * max(1, r.width - 1))
    sy = int(r.top + clamp01(y1) * max(1, r.height - 1))
    ex = int(r.left + clamp01(x2) * max(1, r.width - 1))
    ey = int(r.top + clamp01(y2) * max(1, r.height - 1))

    uia.drag_screen(
        start_x=sx,
        start_y=sy,
        end_x=ex,
        end_y=ey,
        duration_ms=duration_ms,
        steps=steps,
        pre_hold_ms=pre_hold_ms,
        post_hold_ms=post_hold_ms,
        backend=backend,
    )
    console.print(
        f"Dragged inside canvas: ({sx},{sy}) -> ({ex},{ey}) | canvas={r.left},{r.top},{r.right},{r.bottom}"
    )


@app.command("run")
def run_macro_cmd(path: Path = typer.Argument(..., exists=True)):
    """Run a YAML macro."""
    macro.run_macro(path)
    console.print(f"Done: {path}")


# ---------------------------------------------------------------------------
# Browser automation commands (Playwright)
# ---------------------------------------------------------------------------


@app.command("browser-open")
def browser_open_cmd(
    url: str = typer.Option(..., help="URL to open"),
    browser: str = typer.Option("chromium", help="Browser engine: chromium|firefox|webkit"),
    headless: bool = typer.Option(False, help="Run headless (no visible window)"),
):
    """Open a URL in a persistent browser session (Playwright).

    The session persists between commands (login cookies are saved).
    """
    result = browser_mod.open_url(url, browser=browser, headless=headless)  # type: ignore[arg-type]
    console.print(json.dumps(result, ensure_ascii=False, indent=2))


@app.command("browser-navigate")
def browser_navigate_cmd(
    url: str = typer.Option(..., help="URL to navigate to"),
):
    """Navigate the current browser session to a new URL."""
    result = browser_mod.navigate(url)
    console.print(json.dumps(result, ensure_ascii=False, indent=2))


@app.command("browser-snapshot")
def browser_snapshot_cmd(
    fmt: str = typer.Option("aria", help="Output format: aria|text"),
):
    """Snapshot the current page (accessibility tree or text).

    'aria' = structured accessibility tree (roles, names).
    'text' = raw visible text content.
    """
    result = browser_mod.snapshot(fmt=fmt)
    if fmt == "aria" and isinstance(result.get("content"), dict):
        console.print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        console.print(f"URL: {result['url']}")
        console.print(f"Title: {result['title']}")
        console.print("---")
        console.print(str(result.get("content", "")))


@app.command("browser-click")
def browser_click_cmd(
    selector: Optional[str] = typer.Option(None, help="CSS selector"),
    text: Optional[str] = typer.Option(None, help="Click element by visible text"),
    exact: bool = typer.Option(False, help="Exact text match"),
):
    """Click an element on the current page."""
    result = browser_mod.click(selector=selector, text=text, exact=exact)
    console.print(json.dumps(result, ensure_ascii=False, indent=2))


@app.command("browser-type")
def browser_type_cmd(
    selector: str = typer.Option(..., help="CSS selector for the input"),
    text: str = typer.Option(..., help="Text to type"),
    no_clear: bool = typer.Option(False, "--no-clear", help="Do not clear field before typing"),
):
    """Type text into an input element on the current page."""
    result = browser_mod.type_text(selector=selector, text=text, clear=not no_clear)
    console.print(json.dumps(result, ensure_ascii=False, indent=2))


@app.command("browser-screenshot")
def browser_screenshot_cmd(
    out: str = typer.Option("screenshot.png", help="Output file path (.png)"),
    full_page: bool = typer.Option(False, help="Capture full scrollable page"),
):
    """Take a screenshot of the current page."""
    result = browser_mod.screenshot(out=out, full_page=full_page)
    console.print(json.dumps(result, ensure_ascii=False, indent=2))


@app.command("browser-eval")
def browser_eval_cmd(
    js: str = typer.Option(..., help="JavaScript expression to evaluate"),
):
    """Evaluate JavaScript on the current page and print the result."""
    result = browser_mod.evaluate(js)
    console.print(json.dumps(result, ensure_ascii=False, indent=2))


@app.command("browser-links")
def browser_links_cmd():
    """List all links on the current page."""
    result = browser_mod.get_links()
    console.print(json.dumps(result, ensure_ascii=False, indent=2))


_HELP_REFERENCE: list[dict] = [
    {
        "category": "Window Management",
        "commands": [
            {
                "name": "list-windows",
                "desc": "List all top-level windows.",
                "options": ["--title-regex REGEX", "--json"],
                "example": 'hf list-windows --title-regex "Notepad"',
            },
            {
                "name": "focus",
                "desc": "Bring a window to the foreground.",
                "options": ["--title TEXT", "--title-regex REGEX", "--handle INT"],
                "example": 'hf focus --title-regex "Notepad"',
            },
            {
                "name": "start",
                "desc": "Launch an app via the Windows Start menu.",
                "options": ["--app TEXT (required)"],
                "example": 'hf start --app "Notepad"',
            },
            {
                "name": "open-path",
                "desc": "Open a folder in File Explorer.",
                "options": ["--path TEXT (required)", "--direct / --no-direct"],
                "example": r'hf open-path --path "C:\Users\lijin\Downloads"',
            },
        ],
    },
    {
        "category": "UI Discovery",
        "commands": [
            {
                "name": "tree",
                "desc": "Dump the UIA element tree for a window (JSON). Use before clicking anything.",
                "options": [
                    "--title / --title-regex / --handle",
                    "--depth INT (default 3)",
                    "--max-nodes INT (default 5000)",
                ],
                "example": 'hf tree --title-regex "Paint" --depth 6',
            },
            {
                "name": "list-controls",
                "desc": "Print a quick table of all controls in a window.",
                "options": ["--title / --title-regex / --handle", "--depth INT", "--limit INT"],
                "example": 'hf list-controls --title-regex "Paint"',
            },
            {
                "name": "inspect",
                "desc": "Inspect the UI element currently under the mouse cursor.",
                "options": ["--json"],
                "example": "hf inspect --json",
            },
            {
                "name": "resolve",
                "desc": "Test whether a selector JSON matches an element in the current UI.",
                "options": ["--selector-file PATH", "--selector-json TEXT", "--title-regex REGEX"],
                "example": "hf resolve --selector-file sel.json --title-regex Notepad",
            },
            {
                "name": "canvas-selector",
                "desc": "Detect the largest drawable canvas/pane in a window and return its selector JSON.",
                "options": ["--title / --title-regex / --handle"],
                "example": 'hf canvas-selector --title-regex "Paint"',
            },
        ],
    },
    {
        "category": "Interaction (Click / Type)",
        "commands": [
            {
                "name": "click",
                "desc": "Click a UIA control identified by name, automation-id, or control type.",
                "options": [
                    "--title / --title-regex / --handle",
                    "--name TEXT",
                    "--name-regex REGEX",
                    "--auto-id TEXT",
                    "--control-type TEXT  (e.g. Button, Hyperlink, Edit)",
                    "--control TEXT  (pywinauto best_match)",
                    "--timeout INT (default 20s)",
                ],
                "example": 'hf click --title-regex "LinkedIn" --name "Sign in" --control-type Hyperlink',
            },
            {
                "name": "type",
                "desc": "Type text into a UIA control (Edit, RichEdit, etc.).",
                "options": [
                    "--title / --title-regex / --handle",
                    "--text TEXT (required)",
                    "--enter / --no-enter",
                    "--name / --auto-id / --control-type / --control",
                    "--timeout INT",
                ],
                "example": 'hf type --title-regex "Notepad" --control Edit --text "Hello!" --enter',
            },
            {
                "name": "click-at",
                "desc": "Click at window-relative pixel coordinates.",
                "options": ["--title / --title-regex / --handle", "--x INT (required)", "--y INT (required)"],
                "example": 'hf click-at --title-regex "Paint" --x 400 --y 300',
            },
        ],
    },
    {
        "category": "Mouse & Drag",
        "commands": [
            {
                "name": "drag",
                "desc": "Drag within a window using window-relative coordinates.",
                "options": [
                    "--title / --title-regex / --handle",
                    "--start-x / --start-y / --end-x / --end-y (all required)",
                    "--duration-ms INT",
                    "--steps INT",
                ],
                "example": 'hf drag --title-regex "Paint" --start-x 100 --start-y 100 --end-x 400 --end-y 400',
            },
            {
                "name": "drag-screen",
                "desc": "Drag using absolute screen pixel coordinates.",
                "options": [
                    "--start-x / --start-y / --end-x / --end-y (all required)",
                    "--backend pywinauto|sendinput (default pywinauto)",
                    "--duration-ms / --steps / --pre-hold-ms / --post-hold-ms",
                ],
                "example": "hf drag-screen --start-x 350 --start-y 320 --end-x 950 --end-y 620 --backend sendinput",
            },
            {
                "name": "drag-canvas",
                "desc": "Drag inside the auto-detected canvas/content area. Preferred for drawing apps.",
                "options": [
                    "--title / --title-regex / --handle",
                    "--x1 / --y1 / --x2 / --y2 (0.0–1.0 fractions of canvas size)",
                    "--backend sendinput|pywinauto",
                    "--pad INT (canvas edge padding)",
                ],
                "example": 'hf drag-canvas --title-regex "Paint" --x1 0.2 --y1 0.2 --x2 0.7 --y2 0.7 --backend sendinput',
            },
        ],
    },
    {
        "category": "Record & Replay",
        "commands": [
            {
                "name": "record",
                "desc": "Record a UI macro (interactive or passive). Saves to YAML.",
                "options": [
                    "--out PATH (default macro.yaml)",
                    "--passive  (auto-capture clicks/keys; press F9 to stop)",
                    "--verbose",
                    "--window-title-regex REGEX",
                ],
                "example": "hf record --out my_flow.yaml --passive",
            },
            {
                "name": "run",
                "desc": "Execute a recorded YAML macro.",
                "options": ["PATH (positional, required)"],
                "example": "hf run my_flow.yaml",
            },
        ],
    },
    {
        "category": "Browser (Playwright)",
        "commands": [
            {
                "name": "browser-open",
                "desc": "Open a URL in a persistent browser session (login cookies survive).",
                "options": [
                    "--url TEXT (required)",
                    "--browser chromium|firefox|webkit",
                    "--headless / --no-headless",
                ],
                "example": 'hf browser-open --url "https://github.com"',
            },
            {
                "name": "browser-navigate",
                "desc": "Navigate the active browser session to a new URL.",
                "options": ["--url TEXT (required)"],
                "example": 'hf browser-navigate --url "https://github.com/login"',
            },
            {
                "name": "browser-snapshot",
                "desc": "Get the current page as an accessibility tree (aria) or plain text.",
                "options": ["--fmt aria|text (default aria)"],
                "example": "hf browser-snapshot --fmt aria",
            },
            {
                "name": "browser-click",
                "desc": "Click an element on the current page by CSS selector or visible text.",
                "options": ["--selector CSS", "--text TEXT", "--exact / --no-exact"],
                "example": 'hf browser-click --text "Sign in"',
            },
            {
                "name": "browser-type",
                "desc": "Type into an input on the current page.",
                "options": ["--selector CSS (required)", "--text TEXT (required)", "--no-clear"],
                "example": 'hf browser-type --selector "#email" --text "user@example.com"',
            },
            {
                "name": "browser-screenshot",
                "desc": "Save a screenshot of the current page.",
                "options": ["--out PATH (default screenshot.png)", "--full-page"],
                "example": "hf browser-screenshot --out page.png --full-page",
            },
            {
                "name": "browser-eval",
                "desc": "Evaluate JavaScript in the current page context.",
                "options": ["--js TEXT (required)"],
                "example": 'hf browser-eval --js "document.title"',
            },
            {
                "name": "browser-links",
                "desc": "List all hyperlinks on the current page (JSON).",
                "options": [],
                "example": "hf browser-links",
            },
        ],
    },
]


@app.command("help")
def agent_help(
    json_out: bool = typer.Option(False, "--json", help="Output structured JSON (machine-readable)"),
    category: Optional[str] = typer.Option(None, "--category", "-c", help="Filter to a specific category"),
):
    """Agent-friendly command reference. Lists all commands grouped by category with key options and examples.

    Use --json for machine-readable output. Use --category to filter (e.g. 'browser', 'drag').
    """
    data = _HELP_REFERENCE
    if category:
        cat_lower = category.lower()
        data = [g for g in data if cat_lower in g["category"].lower()]
        if not data:
            console.print(f"[yellow]No category matching '{category}'. Available:[/yellow]")
            for g in _HELP_REFERENCE:
                console.print(f"  • {g['category']}")
            raise typer.Exit(1)

    if json_out:
        console.print(json.dumps(data, ensure_ascii=False, indent=2))
        return

    from rich.panel import Panel
    from rich.text import Text

    console.print()
    console.print(
        Panel(
            "[bold cyan]handsfree-windows[/bold cyan] — agent command reference\n"
            "[dim]Run [bold]hf <command> --help[/bold] for full option details.[/dim]\n"
            "[dim]Run [bold]hf help --json[/bold] for machine-readable output.[/dim]",
            expand=False,
        )
    )

    for group in data:
        console.print(f"\n[bold yellow]{group['category']}[/bold yellow]")
        t = Table(show_header=True, header_style="bold", box=None, pad_edge=False, show_edge=False)
        t.add_column("Command", style="bold green", min_width=18, no_wrap=True)
        t.add_column("Description", style="white", min_width=40)
        t.add_column("Key Options / Example", style="dim")
        for cmd in group["commands"]:
            opts = "\n".join(cmd["options"]) if cmd["options"] else ""
            detail = (opts + "\n[bold]eg:[/bold] " + cmd["example"]).strip()
            t.add_row(cmd["name"], cmd["desc"], detail)
        console.print(t)

    console.print()


if __name__ == "__main__":
    app()
