# k5enru — EN↔RU Keyboard Layout Converter

Select text typed in the wrong layout and press the hotkey to convert it. If nothing is selected, the keyboard layout is toggled.

## Usage

- Select mistyped text
- Press the hotkey (default: `F16`)
- Text is converted and layout switches automatically
- If nothing is selected — layout switches without converting

## Configuration

Edit `config.ini` next to the exe:

```ini
[settings]
hotkey = F16
```

## Run on Windows Startup (Task Scheduler)

1. Open **Task Scheduler** → **Create Task** (not "Basic Task")
2. **General** tab:
   - Name: `k5enru`
   - Check **"Run with highest privileges"**
3. **Triggers** tab → New:
   - Begin the task: **At log on**
4. **Actions** tab → New:
   - Program/script: full path to `k5enru.exe`, e.g. `C:\Programs\k5enru\k5enru.exe`
5. **Conditions** tab:
   - Uncheck **"Start the task only if the computer is on AC power"**
6. Click **OK**

> Running with highest privileges ensures the hotkey works in all apps, including those running as Administrator.

## Files

| File | Description |
|------|-------------|
| `k5enru.exe` | Main executable |
| `config.ini` | Hotkey configuration |
| `icon.png` | Tray icon (can be replaced) |
