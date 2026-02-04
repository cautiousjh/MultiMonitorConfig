# Multi Monitor Config

A Windows monitor configuration management tool. Quickly switch between saved monitor profiles from the system tray.

## Features

- **Save/Load Profiles**: Save current monitor settings as profiles and restore with one click
- **Enable/Disable Monitors**: Save profiles with specific monitors disabled
- **Auto-detect**: Automatically re-enable disabled monitors when applying profiles
- **System Tray**: Quick profile switching from tray icon
- **Export/Import**: Backup and share profiles as JSON files

## Installation

### Option 1: Run with Python

```bash
# Install dependencies
pip install -r requirements.txt

# Run
python main.py
# or
run.bat
```

### Option 2: Use standalone exe

Download `MultiMonitorConfig.exe` from the [Releases](../../releases) page

## Usage

1. Double-click tray icon → Opens settings window
2. Click **Save** to save current monitor configuration
3. Select a profile and click **Apply** (or double-click) to restore
4. Right-click tray icon for quick profile switching

### Disabling Monitors

- Uncheck monitors in "Current Monitors" before saving → saves them as disabled
- Check "Disable monitors not in profile" → automatically disables extra monitors when applying

## Profile Storage

Profiles are saved at `%AppData%\MultiMonitorConfig\profiles.json`

## Building

```bash
# Install PyInstaller
pip install pyinstaller

# Build exe
build.bat
```

Output: `dist/MultiMonitorConfig.exe`

## Tech Stack

- Python 3.7+
- customtkinter (UI)
- pystray (System tray)
- Windows API (EnumDisplayDevices, ChangeDisplaySettingsEx, SetDisplayConfig)

## Screenshot

```
┌─────────────────────────────────────┐
│ Current Monitors                    │
│   [✓] \\.\DISPLAY1: 2560x1440 @ 59Hz│
│   [✓] \\.\DISPLAY2: 1920x1080 @ 60Hz│
├─────────────────────────────────────┤
│ Saved Profiles                      │
│   ┌─────────────────────────────┐   │
│   │ Dual Setup                  │   │
│   │ Single Monitor              │   │
│   │ Portable                    │   │
│   └─────────────────────────────┘   │
├─────────────────────────────────────┤
│ Profile Detail                      │
│   [enabled] \\.\DISPLAY1: 2560x1440 │
├─────────────────────────────────────┤
│ [✓] Disable monitors not in profile │
├─────────────────────────────────────┤
│ [Save][Apply][Rename][Delete] ↑↓ ↻  │
└─────────────────────────────────────┘
```

## License

MIT
