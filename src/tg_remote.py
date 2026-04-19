"""
tg_remote.py — Telegram Remote Control for PyMacro
Runs alongside PyMacro (launched by run.bat).
Reads TG credentials from the same userSettings.json.
Uses long-polling to listen for commands from your phone.

Commands:
  /screenshot     — Takes a full-screen screenshot and sends it
  /battery        — Sends current battery % and charging status
  /batterySaveOn  — Start monitoring battery (alert at 100% and 20%)
  /batterySaveOff — Stop battery monitoring
  /reset          — Stop current playback and restart it
  /shutdown       — Shuts down the laptop (with 10 second delay)
"""

import json
import os
import sys
import time
import tempfile
import threading
import urllib.request
import urllib.parse
import urllib.error
from typing import Optional

# ── Constants ─────────────────────────────────────────────────────────────────
POLL_TIMEOUT      = 30   # seconds for long-poll
BATTERY_CHECK     = 60   # seconds between battery checks
SIGNAL_RESET      = os.path.join(tempfile.gettempdir(), "pymacro_reset.signal")
SIGNAL_STOP       = os.path.join(tempfile.gettempdir(), "pymacro_stop.signal")
SIGNAL_SETLIMIT   = os.path.join(tempfile.gettempdir(), "pymacro_setlimit.signal")
SIGNAL_START      = os.path.join(tempfile.gettempdir(), "pymacro_start.signal")
SIGNAL_QUIT       = os.path.join(tempfile.gettempdir(), "pymacro_quit.signal")
LOG_FILE          = os.path.join(os.path.dirname(__file__), "tg_remote.log")

# Settings path (mirrors UserSettings logic)
if sys.platform == "win32":
    _SETTINGS_DIR = os.path.join(os.getenv("LOCALAPPDATA", ""), "PyMacroRecord")
else:
    _SETTINGS_DIR = os.path.join(os.path.expanduser("~"), ".config", "PyMacroRecord")
SETTINGS_FILE = os.path.join(_SETTINGS_DIR, "userSettings.json")


# ── Logging ───────────────────────────────────────────────────────────────────
def log(msg: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


# ── Credentials ───────────────────────────────────────────────────────────────
def load_credentials():
    """Read token and chat_id from userSettings.json."""
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            settings = json.load(f)
        others = settings.get("Others", {})
        token   = others.get("Telegram_Token", "").strip()
        chat_id = others.get("Telegram_Chat_ID", "").strip()
        return token, chat_id
    except Exception as e:
        log(f"ERROR reading settings: {e}")
        return "", ""


def get_battery_save_state() -> bool:
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            settings = json.load(f)
        return bool(settings.get("Others", {}).get("Battery_Save_Active", False))
    except Exception:
        return False


def set_battery_save_state(state: bool):
    try:
        if not os.path.exists(SETTINGS_FILE):
            return
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            settings = json.load(f)
        if "Others" not in settings:
            settings["Others"] = {}
        settings["Others"]["Battery_Save_Active"] = state
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        log(f"Failed to save battery state: {e}")


# ── Telegram API helpers ───────────────────────────────────────────────────────
BASE = "https://api.telegram.org/bot"

def tg_get(token: str, method: str, params: Optional[dict] = None, timeout: int = 35):
    url = f"{BASE}{token}/{method}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    try:
        with urllib.request.urlopen(url, timeout=timeout) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        log(f"HTTP {e.code} on {method}: {e.read().decode()[:200]}")
    except Exception as e:
        log(f"GET {method} error: {e}")
    return None


def tg_post(token: str, method: str, data: dict, timeout: int = 15):
    url = f"{BASE}{token}/{method}"
    encoded = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=encoded)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode())
    except Exception as e:
        log(f"POST {method} error: {e}")
    return None


def send_text(token: str, chat_id: str, text: str, append_help: bool = True):
    """Send a text message. Appends /help by default so the user can tap it to re-open the menu."""
    if append_help:
        text = text + "\n\n/help"
    tg_post(token, "sendMessage", {"chat_id": chat_id, "text": text})


def send_photo(token: str, chat_id: str, image_bytes: bytes, filename: str = "screenshot.png"):
    """Send a photo using multipart form upload."""
    import uuid
    boundary = uuid.uuid4().hex
    CRLF = b"\r\n"

    def field(name, value):
        return (
            f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n{value}\r\n'
            .encode()
        )

    body = (
        field("chat_id", chat_id)
        + f'--{boundary}\r\nContent-Disposition: form-data; name="photo"; filename="{filename}"\r\nContent-Type: image/png\r\n\r\n'.encode()
        + image_bytes
        + CRLF
        + f"--{boundary}--\r\n".encode()
    )

    url = f"{BASE}{token}/sendPhoto"
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode())
    except Exception as e:
        log(f"send_photo error: {e}")
    return None


# ── Command Handlers ───────────────────────────────────────────────────────────
def cmd_screenshot(token: str, chat_id: str):
    try:
        from PIL import ImageGrab
        import io
        log("Taking screenshot...")
        img = ImageGrab.grab()
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        img_bytes = buf.getvalue()
        result = send_photo(token, chat_id, img_bytes, "screenshot.png")
        if result and result.get("ok"):
            log("Screenshot sent.")
        else:
            send_text(token, chat_id, "❌ Failed to send screenshot.")
            log("Screenshot send failed.")
    except Exception as e:
        send_text(token, chat_id, f"❌ Screenshot error: {e}")
        log(f"Screenshot error: {e}")


def cmd_battery(token: str, chat_id: str):
    try:
        import psutil
        batt = psutil.sensors_battery()
        if batt is None:
            send_text(token, chat_id, "🔋 No battery detected (desktop PC?).")
            return
        pct      = round(batt.percent, 1)
        charging = "⚡ Charging" if batt.power_plugged else "🔌 Not Charging"
        send_text(token, chat_id, f"🔋 Battery: {pct}% | {charging}")
        log(f"Battery reported: {pct}% {charging}")
    except ImportError:
        send_text(token, chat_id, "❌ psutil not installed. Run: pip install psutil")
    except Exception as e:
        send_text(token, chat_id, f"❌ Battery error: {e}")
        log(f"Battery error: {e}")


def cmd_shutdown(token: str, chat_id: str):
    send_text(token, chat_id, "⚠️ Shutting down laptop in 10 seconds...\nRun 'shutdown /a' to cancel.")
    log("Shutdown command received. Executing in 10s.")
    time.sleep(2)  # Give Telegram a moment to send the message
    os.system("shutdown /s /t 10")


def cmd_reset(token: str, chat_id: str):
    """Signal PyMacro to stop and restart playback via a shared temp file."""
    try:
        with open(SIGNAL_RESET, "w") as f:
            f.write("reset")
        send_text(token, chat_id, "🔄 Reset signal sent to PyMacro.")
        log("Reset signal file written.")
    except Exception as e:
        send_text(token, chat_id, f"❌ Reset error: {e}")
        log(f"Reset error: {e}")


def cmd_stop(token: str, chat_id: str):
    """Signal PyMacro to stop the current playback."""
    try:
        with open(SIGNAL_STOP, "w") as f:
            f.write("stop")
        send_text(token, chat_id, "⏹️ Stop signal sent to PyMacro.")
        log("Stop signal file written.")
    except Exception as e:
        send_text(token, chat_id, f"❌ Stop error: {e}")
        log(f"Stop error: {e}")


def cmd_restart(token: str, chat_id: str):
    """Restart the laptop."""
    send_text(token, chat_id, "🔁 Restarting laptop in 10 seconds...\nRun 'shutdown /a' to cancel.")
    log("Restart command received. Executing in 10s.")
    time.sleep(2)
    os.system("shutdown /r /t 10")


# ── /setLimit multi-step conversation ─────────────────────────────────────────
# Tracks the current state of a multi-step conversation with the user.
# States:
#   None                              → no pending conversation
#   {"step": "waiting_for_limit"}     → asked user for a number
#   {"step": "waiting_for_confirm",
#    "limit": N}                      → asked user yes/no to start playback
_pending: dict = {}


def cmd_set_limit(token: str, chat_id: str):
    """Begin the /setLimit multi-step conversation."""
    global _pending
    _pending = {"step": "waiting_for_limit"}
    send_text(token, chat_id, "🔢 How many runs do you need?\n(Reply with a number, or /cancel to abort)")
    log("/setLimit started — waiting for number.")


# ── Battery Save Mode ──────────────────────────────────────────────────────────
_battery_save_active = False
_battery_save_thread: Optional[threading.Thread] = None


def _battery_monitor_loop(token: str, chat_id: str):
    """Background thread: check battery every BATTERY_CHECK seconds."""
    global _battery_save_active
    log("Battery Save Mode started.")
    alerted_100 = False
    alerted_20  = False

    try:
        import psutil
    except ImportError:
        send_text(token, chat_id, "❌ psutil not installed. Battery Save Mode cannot start.")
        _battery_save_active = False
        return

    while _battery_save_active:
        try:
            batt = psutil.sensors_battery()
            if batt:
                pct      = batt.percent
                plugged  = batt.power_plugged

                # Alert when fully charged (≥100% while plugged in)
                if plugged and pct >= 100 and not alerted_100:
                    send_text(token, chat_id, "✅ Battery is FULL (100%)! You can unplug the charger. 🔌")
                    log("Battery Save: 100% alert sent.")
                    alerted_100 = True
                    alerted_20  = False  # Reset so 20% alert can fire again later

                # Reset 100% alert when unplugged (so it can fire again next charge cycle)
                if not plugged and alerted_100 and pct < 99:
                    alerted_100 = False

                # Alert when battery drops to 20% while unplugged
                if not plugged and pct <= 20 and not alerted_20:
                    send_text(token, chat_id, f"⚠️ Battery LOW: {round(pct, 1)}%! Plug in the charger now! 🔋")
                    log(f"Battery Save: {pct}% low alert sent.")
                    alerted_20  = True
                    alerted_100 = False  # Reset so 100% alert can fire again next charge

                # Reset 20% alert when plugged back in
                if plugged and alerted_20:
                    alerted_20 = False

        except Exception as e:
            log(f"Battery monitor error: {e}")

        # Sleep in small increments to respond quickly to stop requests
        for _ in range(BATTERY_CHECK):
            if not _battery_save_active:
                break
            time.sleep(1)

    log("Battery Save Mode stopped.")


def cmd_battery_save_on(token: str, chat_id: str):
    global _battery_save_active, _battery_save_thread
    if _battery_save_active:
        send_text(token, chat_id, "ℹ️ Battery Save Mode is already ON.")
        return
    _battery_save_active = True
    set_battery_save_state(True)
    _battery_save_thread = threading.Thread(
        target=_battery_monitor_loop, args=(token, chat_id), daemon=True
    )
    _battery_save_thread.start()
    send_text(token, chat_id, "🔋 Battery Save Mode ON.\nYou'll be notified at 100% (unplug) and 20% (plug in).")
    log("Battery Save Mode enabled.")


def cmd_battery_save_off(token: str, chat_id: str):
    global _battery_save_active
    if not _battery_save_active:
        send_text(token, chat_id, "ℹ️ Battery Save Mode is already OFF.")
        return
    _battery_save_active = False
    set_battery_save_state(False)
    send_text(token, chat_id, "🔕 Battery Save Mode OFF.")
    log("Battery Save Mode disabled.")


# ── Command dispatcher ─────────────────────────────────────────────────────────
COMMANDS = {
    "/screenshot":     cmd_screenshot,
    "/battery":        cmd_battery,
    "/batterysaveon":  cmd_battery_save_on,
    "/batterysaveoff": cmd_battery_save_off,
    "/reset":          cmd_reset,
    "/stop":           cmd_stop,
    "/restartlaptop":  cmd_restart,
    "/setlimit":       cmd_set_limit,
    "/shutdownlaptop": cmd_shutdown,
}

HELP_TEXT = (
    "📋 *PyMacro Remote Commands*\n\n"
    "**PyMacro Commands**\n"
    "/stop — Stop current playback\n"
    "/reset — Stop & restart playback\n"
    "/setLimit — Change the run limit\n\n"
    "**Laptop Commands**\n"
    "/screenshot — Take a screenshot\n"
    "/battery — Check battery status\n"
    "/batterySaveOn — Alert at 100% & 20% battery\n"
    "/batterySaveOff — Stop battery alerts\n"
    "/restartLaptop — Restart the laptop\n"
    "/shutdownLaptop — Shut down the laptop\n\n"
    "/help — Show this message"
)


def _handle_pending(token: str, chat_id: str, text: str) -> bool:
    """Handle a reply in the context of a pending multi-step conversation.
    Returns True if the message was consumed by the pending flow."""
    global _pending
    if not _pending:
        return False

    step = _pending.get("step")

    # cancellation at any step
    if text.strip().lower() in ("/cancel", "cancel"):
        _pending = {}
        send_text(token, chat_id, "❌ Cancelled.")
        return True

    # ── Step 1: waiting for the run count number ──
    if step == "waiting_for_limit":
        try:
            limit = int(text.strip())
            if limit <= 0:
                raise ValueError
        except ValueError:
            send_text(token, chat_id, "⚠️ Please reply with a positive whole number (e.g. 400), or /cancel.")
            return True

        # Write the limit to the signal file — PyMacro will pick it up
        try:
            with open(SIGNAL_SETLIMIT, "w") as f:
                f.write(str(limit))
            log(f"/setLimit: limit {limit} written to signal file.")
        except Exception as e:
            send_text(token, chat_id, f"❌ Failed to write limit signal: {e}")
            _pending = {}
            return True

        _pending = {"step": "waiting_for_confirm", "limit": limit}
        send_text(
            token, chat_id,
            f"✅ Definite limit changed to {limit}!\n"
            f"The script will end after run {limit}.\n\n"
            f"Do you want to start playback now? (yes / no)"
        )
        return True

    # ── Step 2: waiting for yes/no on starting playback ──
    if step == "waiting_for_confirm":
        answer = text.strip().lower()
        limit  = _pending.get("limit", "?")
        _pending = {}

        if answer in ("yes", "y"):
            try:
                with open(SIGNAL_START, "w") as f:
                    f.write("start")
                log(f"/setLimit: start signal written after limit={limit}.")
            except Exception as e:
                send_text(token, chat_id, f"❌ Failed to send start signal: {e}")
                return True
            send_text(token, chat_id, "▶️ Playback starting!")
        elif answer in ("no", "n"):
            send_text(token, chat_id, "👍 Limit set. Playback not started.")
        else:
            send_text(token, chat_id, "⚠️ Please reply yes or no.")
            # Put the state back so they can answer again
            _pending = {"step": "waiting_for_confirm", "limit": limit}
        return True

    return False


def dispatch(token: str, chat_id: str, text: str):
    """Parse and execute a command from a Telegram message."""

    # First: check if we're mid-conversation (e.g. /setLimit flow)
    if _handle_pending(token, chat_id, text):
        return

    # If it's not a command (doesn't start with /), ignore silently
    if not text.startswith("/"):
        return

    # Telegram bot commands can have @botname suffix, strip it
    cmd = text.strip().split()[0].lower().split("@")[0]

    if cmd == "/help":
        send_text(token, chat_id, HELP_TEXT, append_help=False)
        return

    handler = COMMANDS.get(cmd)
    if handler:
        log(f"Executing command: {cmd}")
        threading.Thread(target=handler, args=(token, chat_id), daemon=True).start()
    else:
        send_text(token, chat_id, f"❓ Unknown command: {text.strip()}\nSend /help for a list.")


# ── Main polling loop ──────────────────────────────────────────────────────────
def main():
    # Prevent multiple instances of tg_remote fighting for Telegram updates
    lock_file = os.path.join(tempfile.gettempdir(), "pymacro_tg_remote.lock")
    if os.path.exists(lock_file):
        try:
            with open(lock_file, "r") as f:
                old_pid = int(f.read().strip())
            try:
                os.kill(old_pid, 0)
                sys.exit(0)
            except OSError as e:
                if getattr(e, 'winerror', None) == 5 or e.errno == 13:
                    sys.exit(0)
                pass
        except Exception:
            pass
            
    try:
        with open(lock_file, "w") as f:
            f.write(str(os.getpid()))
    except Exception:
        pass

    log("=" * 50)
    log("tg_remote.py starting up...")

    token, chat_id = load_credentials()

    if not token or not chat_id:
        log("ERROR: No Telegram token or chat ID found in userSettings.json.")
        log("Please configure Telegram in PyMacro settings first, then relaunch.")
        log("tg_remote.py will exit.")
        sys.exit(1)

    log(f"Token loaded (...{token[-6:]}), Chat ID: {chat_id}")
    send_text(token, chat_id, "✅ PyMacro Remote is online.\nSend /help for commands.")
    log("Startup message sent. Entering polling loop...")

    global _battery_save_active, _battery_save_thread
    _battery_save_active = get_battery_save_state()
    if _battery_save_active:
        _battery_save_thread = threading.Thread(
            target=_battery_monitor_loop, args=(token, chat_id), daemon=True
        )
        _battery_save_thread.start()
        log("Restored Battery Save Mode from settings and started thread.")

    offset = None

    while True:
        try:
            # Check if PyMacro closed and left a quit signal
            if os.path.exists(SIGNAL_QUIT):
                try:
                    os.remove(SIGNAL_QUIT)
                except Exception:
                    pass
                log("Quit signal received from PyMacro. Shutting down tg_remote.")
                break

            params = {"timeout": POLL_TIMEOUT, "allowed_updates": ["message"]}
            if offset is not None:
                params["offset"] = offset

            data = tg_get(token, "getUpdates", params, timeout=POLL_TIMEOUT + 5)

            if data and data.get("ok"):
                for update in data.get("result", []):
                    offset = update["update_id"] + 1

                    msg = update.get("message", {})
                    from_chat_id = str(msg.get("chat", {}).get("id", ""))
                    text = msg.get("text", "").strip()

                    # Security: only respond to the configured chat
                    if from_chat_id != str(chat_id):
                        log(f"Ignored message from unknown chat: {from_chat_id}")
                        continue

                    if text:
                        log(f"Received message: {text!r}")
                        dispatch(token, chat_id, text)

        except KeyboardInterrupt:
            log("Interrupted. Shutting down tg_remote.")
            break
        except Exception as e:
            log(f"Polling loop error: {e}")
            time.sleep(5)  # Back off before retrying


if __name__ == "__main__":
    main()
