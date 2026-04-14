# Parsec Monitor

<img width="660" height="534" alt="image" src="https://github.com/user-attachments/assets/6d81b036-c7ee-44a4-94ef-dbf6da119e2e" />

Background Parsec connection monitor with a system tray icon, user whitelist, and software input lock.

## Features

- Tracks Parsec connect/disconnect events via `log.txt`
- System tray icon with notifications
- Event log in the GUI
- **User whitelist** with automatic actions:
  - On connect — release the software input lock
  - On disconnect — engage the software input lock
- **Software lock** (`SoftLock`):
  - Blocks keyboard and mouse via low-level system hooks
  - No visual change (screen stays on)
  - All blocked input attempts are silently logged to `unlock_debug.log`
  - Manual unlock hotkey: **Ctrl + Alt + F13**

## Requirements

- Windows 10/11
- Python 3.10+
- Parsec (installed and launched at least once)

## Installation

```bat
pip install PyQt6
```

## Running

```bat
python parsec_monitor.py
```

Or without a console window:
```bat
pythonw parsec_monitor.py
```

## Files

| File | Description |
|------|-------------|
| `parsec_monitor.py` | Main script |
| `config.json` | Configuration (created automatically) |
| `unlock_debug.log` | Lock/unlock event log |

## Configuration (`config.json`)

```json
{
  "install_type": "per_user",
  "parsec_folder": "C:\\Users\\User\\AppData\\Roaming\\Parsec",
  "whitelist": [
    {
      "parsec_user": "SomeUser#101532125",
      "auto_unlock": true,
      "auto_lock": true
    }
  ]
}
```

### Adding a user to the whitelist

**Via GUI:** click `+ Add` in the right panel of the window.

**Manually:** add an object to `whitelist` with the fields above.

## Hotkeys

| Combination | Action |
|-------------|--------|
| `Ctrl + Alt + F13` | Manually release the software input lock |

## Autostart

Run `install_autostart.bat` as administrator.
To remove from autostart — run `uninstall_autostart.bat`.

## System Tray

- **Double-click** — open main window
- **Right-click → Show** — open main window
- **Right-click → Quit** — exit the program
- Closing the window minimizes to tray, does not exit the program
