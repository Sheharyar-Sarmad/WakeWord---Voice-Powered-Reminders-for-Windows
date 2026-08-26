"""
main.py - WakeWord: Desktop reminders with voice for Windows
Full date/time support + automatic background checker
"""

import json
import time
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
import threading
import winsound
import atexit

# App Identity
APP_NAME = "WakeWord"   # Change this to your preferred app name

# Notification library: plyer
HAS_NOTIFICATION = False
try:
    from plyer import notification
    HAS_NOTIFICATION = True
except ImportError:
    print("ℹ️  Install plyer: pip install plyer")
except Exception:
    pass

class SimpleReminder:
    def __init__(self, data_file="reminders.json"):
        self.data_file = Path(data_file)
        self.reminders = []
        self.checker_running = True
        self.checker_thread = None
        self._load_reminders()
        self.start_checker_background()

    def _load_reminders(self):
        if self.data_file.exists():
            try:
                with open(self.data_file, 'r') as f:
                    self.reminders = json.load(f)
                for r in self.reminders:
                    if 'datetime' not in r and 'time' in r:
                        today = datetime.now().date()
                        dt_str = f"{today.isoformat()} {r['time']}"
                        r['datetime'] = dt_str
                print(f"✅ Loaded {len(self.reminders)} reminders")
            except Exception as e:
                print(f"⚠️  Load error: {e}")
                self.reminders = []
        else:
            self.reminders = []
            self._save_reminders()

    def _save_reminders(self):
        try:
            with open(self.data_file, 'w') as f:
                json.dump(self.reminders, f, indent=2)
            return True
        except Exception as e:
            print(f"❌ Save error: {e}")
            return False

    def add_reminder(self, title, message, dt_str, repeat_type='once',
                     repeat_interval=0, priority='normal'):
        try:
            reminder_dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
        except ValueError:
            print("❌ Invalid datetime! Use: YYYY-MM-DD HH:MM")
            return False

        if repeat_type == 'once' and reminder_dt < datetime.now():
            print("⚠️  This date/time is in the past!")
            confirm = input("Add anyway? (y/n): ")
            if confirm.lower() != 'y':
                return False

        reminder = {
            'id': len(self.reminders) + 1,
            'title': title,
            'message': message,
            'datetime': reminder_dt.isoformat(),
            'repeat_type': repeat_type,
            'repeat_interval': repeat_interval,
            'priority': priority,
            'active': True,
            'created_at': datetime.now().isoformat(),
            'last_triggered': None
        }

        self.reminders.append(reminder)
        self._save_reminders()

        print(f"✅ Reminder added! (ID: {reminder['id']})")
        print(f"   {title} on {reminder_dt.strftime('%Y-%m-%d %H:%M')} ({repeat_type})")
        return True

    def get_due_reminders(self):
        now = datetime.now()
        due = []

        for reminder in self.reminders:
            if not reminder['active']:
                continue

            try:
                reminder_dt = datetime.fromisoformat(reminder['datetime'])
            except:
                continue

            if reminder_dt > now:
                continue

            last = reminder.get('last_triggered')
            if last:
                last_dt = datetime.fromisoformat(last)
                if reminder['repeat_type'] == 'daily':
                    if last_dt.date() == now.date():
                        continue
                elif reminder['repeat_type'] == 'interval':
                    minutes_passed = (now - last_dt).total_seconds() / 60
                    if minutes_passed < reminder['repeat_interval']:
                        continue
                elif reminder['repeat_type'] == 'once':
                    continue
            due.append(reminder)

        return due

    def trigger_reminder(self, reminder):
        reminder['last_triggered'] = datetime.now().isoformat()
        self._save_reminders()

        self._send_notification(reminder['title'], reminder['message'])
        self._speak(reminder['title'], reminder['message'])
        print(f"🔔 {reminder['title']}: {reminder['message']}")

    def _send_notification(self, title, message):
        if HAS_NOTIFICATION:
            try:
                notification.notify(
                    title=f"{APP_NAME}: {title}",
                    message=message,
                    timeout=10,
                    app_name=APP_NAME   # <-- shows "WakeWord" in notification
                )
                print("📨 Notification sent!")
                return
            except Exception as e:
                print(f"⚠️  plyer error: {e}")

        # Fallback: MessageBox
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, message, f"{APP_NAME}: {title}", 0)
            print("📨 MessageBox sent!")
        except:
            winsound.Beep(1000, 500)
            print(f"\n{'='*50}")
            print(f"🔔 {APP_NAME}: {title}")
            print(f"📝 {message}")
            print(f"{'='*50}")

    def _speak(self, title, message):
        try:
            import win32com.client
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Speak(f"{APP_NAME} reminder: {title}. {message}")
            print("🗣️ Voice played!")
        except ImportError:
            winsound.Beep(1000, 300)
            time.sleep(0.1)
            winsound.Beep(1200, 300)
        except:
            pass

    def list_reminders(self, show_all=False):
        if not self.reminders:
            print("📭 No reminders found")
            return

        print(f"\n{'='*90}")
        print(f"{'ID':<4} {'Title':<20} {'Date/Time':<20} {'Repeat':<12} {'Status'}")
        print(f"{'='*90}")

        for r in self.reminders:
            if not show_all and not r['active']:
                continue
            dt = datetime.fromisoformat(r['datetime']).strftime('%Y-%m-%d %H:%M')
            status = "✅ Active" if r['active'] else "❌ Paused"
            repeat = r['repeat_type']
            if repeat == 'interval':
                repeat = f"every {r['repeat_interval']}min"
            print(f"{r['id']:<4} {r['title'][:20]:<20} {dt:<20} {repeat:<12} {status}")
        print(f"{'='*90}")

    def delete_reminder(self, reminder_id):
        for i, r in enumerate(self.reminders):
            if r['id'] == reminder_id:
                del self.reminders[i]
                self._save_reminders()
                print(f"🗑️  Reminder {reminder_id} deleted")
                return True
        print(f"❌ Reminder {reminder_id} not found")
        return False

    def pause_reminder(self, reminder_id):
        for r in self.reminders:
            if r['id'] == reminder_id:
                r['active'] = False
                self._save_reminders()
                print(f"⏸️  Reminder {reminder_id} paused")
                return True
        print(f"❌ Reminder {reminder_id} not found")
        return False

    def resume_reminder(self, reminder_id):
        for r in self.reminders:
            if r['id'] == reminder_id:
                r['active'] = True
                self._save_reminders()
                print(f"▶️  Reminder {reminder_id} resumed")
                return True
        print(f"❌ Reminder {reminder_id} not found")
        return False

    def _checker_loop(self):
        while self.checker_running:
            try:
                due = self.get_due_reminders()
                for reminder in due:
                    self.trigger_reminder(reminder)
                time.sleep(30)
            except Exception as e:
                print(f"⚠️  Checker error: {e}")
                time.sleep(30)

    def start_checker_background(self):
        if self.checker_thread is None or not self.checker_thread.is_alive():
            self.checker_running = True
            self.checker_thread = threading.Thread(target=self._checker_loop, daemon=True)
            self.checker_thread.start()
            print(f"🔄 {APP_NAME} checker started in background.")

    def stop_checker(self):
        self.checker_running = False
        if self.checker_thread:
            self.checker_thread.join(timeout=1)
            print("⏹️  Checker stopped.")

    def run_checker_foreground(self):
        self.checker_running = True
        print(f"\n🔄 {APP_NAME} checker running in foreground...")
        print("⏰ Checking every 30 seconds")
        print("Press Ctrl+C to stop\n")
        try:
            while self.checker_running:
                due = self.get_due_reminders()
                for reminder in due:
                    self.trigger_reminder(reminder)
                time.sleep(30)
        except KeyboardInterrupt:
            print("\n👋 Stopping checker...")
            self.checker_running = False


# Helper & Menu 
def parse_date_input(date_str):
    if date_str.lower() == 'today':
        return datetime.now().date()
    elif date_str.lower() == 'tomorrow':
        return datetime.now().date() + timedelta(days=1)
    else:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            return None

def interactive_menu():
    print(f"🔔 {APP_NAME} – SIMPLE WINDOWS REMINDER SYSTEM")
    print("=" * 50)

    reminder = SimpleReminder()

    while True:
        print("\n" + "=" * 50)
        print(f"📅 {APP_NAME} REMINDER SYSTEM")
        print("=" * 50)
        print("1. Add Reminder")
        print("2. List Reminders")
        print("3. Delete Reminder")
        print("4. Pause/Resume Reminder")
        print("5. Start Checker (foreground)")
        print("6. Stop Checker")
        print("7. Exit")
        print("=" * 50)

        choice = input("Choose option (1-7): ").strip()

        if choice == "1":
            print("\n📝 NEW REMINDER")
            title = input("Title: ").strip()
            message = input("Message: ").strip()

            while True:
                date_input = input("Date (YYYY-MM-DD) or 'today' or 'tomorrow': ").strip()
                date_obj = parse_date_input(date_input)
                if date_obj:
                    break
                print("❌ Invalid date. Use YYYY-MM-DD, 'today', or 'tomorrow'.")

            while True:
                time_str = input("Time (HH:MM, 24-hour): ").strip()
                try:
                    datetime.strptime(time_str, "%H:%M")
                    break
                except ValueError:
                    print("❌ Invalid time. Use HH:MM (e.g., 14:30)")

            dt_str = f"{date_obj.isoformat()} {time_str}"

            print("\nRepeat Type:")
            print("1. Once")
            print("2. Daily")
            print("3. Interval (every X minutes)")
            repeat_choice = input("Choose (1-3): ").strip()

            repeat_map = {'1': 'once', '2': 'daily', '3': 'interval'}
            repeat_type = repeat_map.get(repeat_choice, 'once')
            interval = 0
            if repeat_type == 'interval':
                interval = int(input("Interval in minutes: "))

            priority = input("Priority (low/normal/high) [normal]: ").strip() or 'normal'

            reminder.add_reminder(title, message, dt_str, repeat_type, interval, priority)

        elif choice == "2":
            show_all = input("Show all (including paused)? (y/n): ").lower() == 'y'
            reminder.list_reminders(show_all)

        elif choice == "3":
            reminder.list_reminders()
            rid = input("Enter reminder ID to delete: ").strip()
            if rid.isdigit():
                reminder.delete_reminder(int(rid))

        elif choice == "4":
            reminder.list_reminders()
            rid = input("Enter reminder ID: ").strip()
            if rid.isdigit():
                action = input("Pause or Resume? (p/r): ").lower()
                if action == 'p':
                    reminder.pause_reminder(int(rid))
                elif action == 'r':
                    reminder.resume_reminder(int(rid))

        elif choice == "5":
            reminder.stop_checker()
            reminder.run_checker_foreground()

        elif choice == "6":
            reminder.stop_checker()

        elif choice == "7":
            print(f"👋 Goodbye from {APP_NAME}!")
            reminder.stop_checker()
            break

        else:
            print("❌ Invalid choice")

def quick_start_examples():
    r = SimpleReminder()
    now = datetime.now()
    examples = [
        {'title': 'Morning Meeting', 'message': 'Daily standup with team', 'datetime': now.replace(hour=9, minute=30).isoformat(), 'repeat': 'daily', 'priority': 'high'},
        {'title': 'Lunch Time', 'message': "Don't forget to eat!", 'datetime': now.replace(hour=13, minute=0).isoformat(), 'repeat': 'daily', 'priority': 'normal'},
        {'title': 'Water Break', 'message': 'Drink water 💧', 'datetime': now.isoformat(), 'repeat': 'interval', 'interval': 60, 'priority': 'low'},
        {'title': 'End of Day', 'message': 'Time to wrap up work', 'datetime': now.replace(hour=18, minute=0).isoformat(), 'repeat': 'daily', 'priority': 'normal'}
    ]
    for ex in examples:
        r.add_reminder(
            ex['title'],
            ex['message'],
            ex['datetime'],
            ex['repeat'],
            ex.get('interval', 0),
            ex.get('priority', 'normal')
        )
    print("\n✅ Example reminders added!")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == '--examples':
            quick_start_examples()
        elif sys.argv[1] == '--checker':
            r = SimpleReminder()
            r.run_checker_foreground()
        else:
            print("Usage:")
            print("  python main.py               # Interactive menu (checker auto-starts)")
            print("  python main.py --examples    # Add example reminders")
            print("  python main.py --checker     # Run checker only (foreground)")
    else:
        interactive_menu()