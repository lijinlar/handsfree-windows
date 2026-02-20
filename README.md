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

Launch an app from Start menu (works for almost anything installed):
```powershell
hf start --app "Notepad"
```

Focus a window (by regex title):
```powershell
hf focus --title-regex "Notepad"
```

List controls for a window:
```powershell
hf list-controls --title-regex "Notepad" --depth 3
```

Type into a control (best-match or auto-id):
```powershell
hf type --title-regex "Notepad" --text "Hello" --control "Edit"
```

Inspect element under cursor (prints a robust selector/path you can paste into a skill):
```powershell
hf inspect --json
```

Record a macro interactively (MVP):
```powershell
hf record --out demo.yaml
hf run demo.yaml
```

## Notes
- Backend is `pywinauto` (UIA). Some apps require elevated privileges.
- This project avoids stealth/botting features; it’s for local automation and accessibility-style control.
