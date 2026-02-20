# handsfree-windows

CLI to control Windows applications using UI Automation (UIA) (and optionally browser automation).

## Goals
- Operate native Windows apps “like a human”: focus windows, find UI controls, click, type, wait.
- Provide a scriptable CLI + a small macro language for repeatable flows.

## Install (dev)
```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -e .
```

## Quick start
List top-level windows:
```powershell
hf list-windows
```

Focus a window (by regex title):
```powershell
hf focus --title "Notepad"
```

List controls for a window:
```powershell
hf list-controls --title "Notepad" --depth 3
```

Type into a control (best-match or auto-id):
```powershell
hf type --title "Notepad" --text "Hello" --control "Edit"
```

## Notes
- Backend is `pywinauto` (UIA). Some apps require elevated privileges.
- This project avoids stealth/botting features; it’s for local automation and accessibility-style control.
