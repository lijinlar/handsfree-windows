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

### Install (dev)
```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -e .
```

### List windows
```powershell
hf list-windows
hf list-windows --json
hf list-windows --title-regex "Outlook"
```

### Launch apps
Launch from Start menu (generic):
```powershell
hf start --app "Outlook"
hf start --app "Paint"
hf start --app "LinkedIn"
```

Open a folder in File Explorer (most reliable = direct explorer.exe):
```powershell
hf open-path --path "C:\\Users\\lijin\\.openclaw\\workspace"
# Human-style fallback:
# hf open-path --path "C:\\...\\workspace" --direct false
```

### Focus a window
```powershell
hf focus --title "LinkedIn"
# or
hf focus --title-regex "Untitled - Paint"
```

### Discover UI (no guessing)
Export the UI Automation tree (JSON):
```powershell
hf tree --title-regex "Untitled - Paint" --depth 10 --max-nodes 30000
```

List controls in a table (quick glance):
```powershell
hf list-controls --title-regex "Untitled - Paint" --depth 4
```

Inspect element under cursor (prints robust selector JSON):
```powershell
hf inspect --json
```

Resolve a selector JSON against the current UI (debugging):
```powershell
hf resolve --selector-file selector.json
# or override the window matcher:
hf resolve --selector-file selector.json --title-regex "Untitled - Paint"
```

### Click and type (by UIA properties)
```powershell
hf click --title "LinkedIn" --name "Sign in with browser" --control-type "Hyperlink"

hf type --title-regex "Notepad" --control "Edit" --text "Hello from handsfree-windows!"
```

### Mouse primitives
Absolute screen drag:
```powershell
hf drag-screen --start-x 350 --start-y 320 --end-x 950 --end-y 620 --backend sendinput
```

Drag inside the detected canvas/content area (avoids out-of-canvas drags):
```powershell
hf drag-canvas --title-regex "Untitled - Paint" --backend sendinput
# tweak shape size/position with fractions:
# hf drag-canvas --title-regex "Untitled - Paint" --x1 0.30 --y1 0.30 --x2 0.55 --y2 0.50
```

### Record + replay (generic macros)
```powershell
hf record --out demo.yaml
hf run demo.yaml
```

## Notes
- Backend is `pywinauto` (UIA). Some apps require elevated privileges.
- This project avoids stealth/botting features; it’s for local automation and accessibility-style control.
