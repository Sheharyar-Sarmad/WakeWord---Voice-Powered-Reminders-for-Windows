# WakeWord 🔔

A lightweight desktop reminder system for Windows with voice and notification support.

## Features

- 📅 Add reminders with custom date and time
- 🔁 Repeat types: once, daily, or custom interval (minutes)
- 🔔 Native Windows notifications (via `plyer`, with `MessageBox` fallback)
- 🗣️ Voice announcements using SAPI text-to-speech
- ⚙️ Background checker thread that runs automatically
- 💾 Persistent storage using a local JSON file
- 🖥️ Simple interactive command-line menu

## Requirements

- Python 3.8+
- Windows OS (uses `winsound`, `ctypes`, and optionally `win32com.client`)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/<your-username>/wakeword.git
   cd wakeword
   ```

2. Install dependencies:
   ```bash
   pip install plyer pywin32
   ```

## Usage

Run the interactive menu (checker starts automatically in the background):

```bash
python main.py
```

Other run modes:

```bash
python main.py --examples   # Add example reminders
python main.py --checker    # Run checker only (foreground)
```

### Menu Options

| Option | Description |
|--------|-------------|
| 1 | Add a new reminder |
| 2 | List all reminders |
| 3 | Delete a reminder |
| 4 | Pause / resume a reminder |
| 5 | Start checker in foreground |
| 6 | Stop background checker |
| 7 | Exit |

## Troubleshooting

- **No notifications appearing** — Ensure `plyer` is installed (`pip install plyer`). If it fails, the app falls back to a Windows `MessageBox` popup automatically.
- **No voice playback** — Install `pywin32` (`pip install pywin32`) to enable SAPI text-to-speech. Without it, the app falls back to audio beeps.
- **Reminders not triggering** — Confirm the background checker is running; it checks every 30 seconds. Use option `5` in the menu to run it in the foreground for debugging.
- **Invalid date/time errors** — Dates must be `YYYY-MM-DD` and times must be `HH:MM` in 24-hour format.

## Configuration

- App name can be changed via the `APP_NAME` variable in `main.py`.
- Reminder data is stored in `reminders.json` in the project root.

## Roadmap

- [ ] GUI interface
- [ ] Cross-platform notification support
- [ ] Recurring reminders on specific weekdays

## License

This project is licensed under the MIT License.

## Author

**Sheharyar Sarmad**
[LinkedIn](https://www.linkedin.com/in/sheharyar-sarmad-9b7736289/)

## Social

**Linkedin**
[Post Link](https://www.linkedin.com/feed/update/urn:li:ugcPost:7498322302254907392/)
