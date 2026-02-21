# Task: Implement Passive Macro Recording

## What to build: `hf record --passive` flag

### Behaviour
- Start recording immediately — no prompts between steps
- **Left click** → resolve UIA element at cursor position using existing `uia.selector_for_element()` + `uia.element_from_point()` → record as `click` step with `selector_candidates`
- **Keyboard typing** → accumulate printable chars into a buffer → flush as a `type` step when:
  - Mouse clicks (focus change)
  - Enter key pressed (record with `enter: true`)
  - 1.5s idle timeout
- **Stop hotkey: F9** → stop recording, save YAML, print path
- **Do NOT record** the stop-hotkey itself

### Implementation plan
1. Add `--passive` flag to the existing `record_macro` command in `cli.py`
2. Create `passive_record()` function in `macro.py` (or a new `recorder.py` file)
3. Use `pynput.mouse.Listener` + `pynput.keyboard.Listener` for global hooks
4. Reuse existing UIA helpers: `uia.element_from_point(x, y)` and `uia.selector_for_element(elem)` to capture selectors
5. Thread-safe: hooks run in threads, main thread waits for stop signal
6. On stop: flush any pending type buffer → save YAML using same format as existing record command

### YAML output format (must match existing format exactly)
```yaml
- action: click
  args:
    selector_candidates:
      - window:
          title_regex: .*Notepad.*
          class_name: Notepad
        targets:
          - auto_id: '15'
            control_type: Edit
            name: ''
    timeout: 20

- action: type
  args:
    selector_candidates:
      - ...same as last click...
    text: Hello world
    enter: false
    timeout: 20
```

### Edge cases to handle
- UIA lookup may fail → skip step + print warning, don't crash
- Rapid clicks (double-click) → record both
- Non-printable keys (arrows, backspace, delete) → flush type buffer, ignore key itself
- Window focus changes → flush type buffer

### pyproject.toml
Add `pynput` to dependencies (already pip-installed).

### Sanity check after implementation
Run: `python -c "from handsfree_windows import macro; print('ok')"`

### When done
Run this EXACT command to notify:
```
openclaw system event --text "Done: passive recording implemented. hf record --passive captures clicks+typing, stops on F9." --mode now
```
