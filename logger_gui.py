# logger_gui.py
import tkinter as tk
import threading
import subprocess
import os
import time
import sys
import ctypes

import logger_bot

def resource_path(relative_path):
    """Get absolute path to resource (works in dev + PyInstaller EXE)."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def append_console(text: str):
    console.configure(state="normal")
    console.insert("end", text.rstrip() + "\n")
    console.see("end")
    console.configure(state="disabled")


def is_setup_message(msg: str) -> bool:
    return msg.strip().startswith("SETUP:")


def clean_setup(msg: str) -> str:
    return msg.strip().removeprefix("SETUP:").strip()


def set_start_enabled(enabled: bool):
    # Thread-safe UI update
    def apply():
        start_btn.config(state="normal" if enabled else "disabled")
    root.after(0, apply)


def open_logs():
    path = logger_bot.get_logs_dir()
    os.makedirs(path, exist_ok=True)
    subprocess.Popen(f'explorer "{path}"')


def open_config_file():
    config_path = getattr(logger_bot, "CONFIG_PATH", os.path.join(os.getcwd(), "config.json"))
    try:
        # If it doesn't exist yet, attempt to trigger creation (logger_bot does this)
        if not os.path.exists(config_path):
            try:
                logger_bot.start_bot_background()
            except Exception:
                pass

        if os.path.exists(config_path):
            os.startfile(config_path)  # Windows-only
            append_console(f"Opened config: {config_path}")
        else:
            append_console("SETUP → config.json not found yet. Click Start again.")
    except Exception as e:
        append_console(f"ERROR: Couldn't open config.json: {e}")


def on_status_update(status: dict):
    # Called from logger_bot threads; marshal to tkinter thread
    def apply():
        bot_state_var.set(f"Bot: {status.get('bot_state', 'UNKNOWN')}")
        logging_state_var.set(f"Logging: {status.get('logging_state', 'OFF')}")
        msgs_var.set(f"Messages logged (this session): {status.get('messages_logged', 0)}")

        last_time = status.get("last_message_time")
        last_var.set(f"Last activity: {last_time if last_time else '—'}")

        err = status.get("last_error")
        error_var.set(f"Error: {err if err else '—'}")

        note = status.get("note")
        if note:
            append_console(note)

        # Keep checkbox in sync with config
        autostart_val = bool(status.get("config_autostart", False))
        if auto_var.get() != autostart_val:
            auto_var.set(autostart_val)

    root.after(0, apply)


logger_bot.subscribe_status(on_status_update)


def start_clicked():
    def work():
        try:
            set_start_enabled(False)
            append_console("Starting…")

            # Boot bot (creates config if missing)
            logger_bot.start_bot_background()

            # Wait up to ~5 seconds for bot to come online
            for _ in range(25):
                if logger_bot.is_bot_running():
                    break
                time.sleep(0.2)

            msg = logger_bot.start_logging()

            if is_setup_message(msg):
                append_console(f"SETUP → {clean_setup(msg)}")
                append_console("SETUP → Opening config.json for you…")
                root.after(0, open_config_file)
            else:
                append_console(msg)

        except Exception as e:
            msg = str(e).strip()
            if is_setup_message(msg):
                append_console(f"SETUP → {clean_setup(msg)}")
                append_console("SETUP → Opening config.json…")
                root.after(0, open_config_file)
            else:
                append_console(f"ERROR: {msg}")

        finally:
            set_start_enabled(True)

    threading.Thread(target=work, daemon=True).start()


def stop_clicked():
    try:
        msg = logger_bot.stop_logging()
        append_console(msg)
    except Exception as e:
        append_console(f"ERROR: {e}")


def toggle_autostart():
    try:
        logger_bot.set_autostart(auto_var.get())
        append_console(f"Auto-start set to {auto_var.get()}.")
    except Exception as e:
        append_console(f"ERROR setting autostart: {e}")


# ---------------- GUI ----------------
myappid = "discord.channel.logger.1.0"
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

root = tk.Tk()

icon_path = resource_path("icon.ico")
try:
    root.iconbitmap(icon_path)
except Exception as e:
    append_console(f"ERROR: Couldn't set window icon: {e}")

root.title("Discord Channel Logger")
root.geometry("520x440")
root.resizable(False, False)

tk.Label(root, text="Discord Channel Logger", font=("Segoe UI", 14, "bold")).pack(pady=10)

# Buttons row
top = tk.Frame(root)
top.pack(pady=4)

start_btn = tk.Button(top, text="Start", width=16, height=2, command=start_clicked)
start_btn.grid(row=0, column=0, padx=6, pady=4)

tk.Button(top, text="Stop", width=16, height=2, command=stop_clicked).grid(row=0, column=1, padx=6, pady=4)
tk.Button(top, text="Open Log Location", width=18, height=2, command=open_logs).grid(row=0, column=2, padx=6, pady=4)

# Auto-start checkbox
auto_var = tk.BooleanVar(value=logger_bot.get_status().get("config_autostart", False))
tk.Checkbutton(
    root,
    text="Auto-start logging when app opens",
    variable=auto_var,
    command=toggle_autostart
).pack(pady=6)

# Status area
status_frame = tk.LabelFrame(root, text="Status", padx=10, pady=8)
status_frame.pack(fill="x", padx=12, pady=6)

bot_state_var = tk.StringVar(value="Bot: OFFLINE")
logging_state_var = tk.StringVar(value="Logging: OFF")
msgs_var = tk.StringVar(value="Messages logged (this session): 0")
last_var = tk.StringVar(value="Last activity: —")
error_var = tk.StringVar(value="Error: —")

tk.Label(status_frame, textvariable=bot_state_var, anchor="w").pack(fill="x")
tk.Label(status_frame, textvariable=logging_state_var, anchor="w").pack(fill="x")
tk.Label(status_frame, textvariable=msgs_var, anchor="w").pack(fill="x")
tk.Label(status_frame, textvariable=last_var, anchor="w").pack(fill="x")
tk.Label(status_frame, textvariable=error_var, anchor="w").pack(fill="x")

# Activity log
console_frame = tk.LabelFrame(root, text="Activity Log", padx=8, pady=8)
console_frame.pack(fill="both", expand=True, padx=12, pady=8)

console = tk.Text(console_frame, height=10, wrap="word", state="disabled")
console.pack(fill="both", expand=True)

append_console("App started.")
append_console("Click Start to connect + begin logging.")
append_console("If config.json is missing/placeholder, it will be created and opened automatically.")


# Auto-start behavior
def do_autostart_if_enabled():
    st = logger_bot.get_status()
    if st.get("config_autostart"):
        append_console("Auto-start enabled. Starting…")
        start_clicked()

root.after(400, do_autostart_if_enabled)

root.mainloop()