from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table

from . import macro, uia

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


@app.command("run")
def run_macro_cmd(path: Path = typer.Argument(..., exists=True)):
    """Run a YAML macro."""
    macro.run_macro(path)
    console.print(f"Done: {path}")


if __name__ == "__main__":
    app()
