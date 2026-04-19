import json
import threading
import urllib.request
import urllib.parse
from tkinter import messagebox

class TelegramNotifier:
    def __init__(self, main_app):
        self.main_app = main_app
        self.base_url = "https://api.telegram.org/bot"

    def get_token(self):
        return self.main_app.settings.settings_dict.get("Others", {}).get("Telegram_Token", "")

    def get_chat_id(self):
        return self.main_app.settings.settings_dict.get("Others", {}).get("Telegram_Chat_ID", "")

    def is_enabled(self):
        token = self.get_token()
        chat_id = self.get_chat_id()
        return bool(token and chat_id)

    def send_message(self, message):
        """Send a message to the configured Telegram chat (Async)"""
        token = self.get_token()
        chat_id = self.get_chat_id()
        
        if not token or not chat_id:
            return False

        def _send():
            try:
                url = f"{self.base_url}{token}/sendMessage"
                data = urllib.parse.urlencode({"chat_id": chat_id, "text": message}).encode()
                req = urllib.request.Request(url, data=data)
                with urllib.request.urlopen(req, timeout=10) as response:
                    pass
            except Exception as e:
                print(f"TG Send Error: {e}")

        threading.Thread(target=_send, daemon=True).start()
        return True

    def test_connection(self, token, chat_id):
        """Blocking call to test connection and validity"""
        try:
            url = f"{self.base_url}{token}/getMe"
            with urllib.request.urlopen(url, timeout=10) as response:
                if response.getcode() == 200:
                    if chat_id:
                        send_url = f"{self.base_url}{token}/sendMessage"
                        data = urllib.parse.urlencode({"chat_id": chat_id, "text": "Test from PyMacro"}).encode()
                        req = urllib.request.Request(send_url, data=data)
                        try:
                            urllib.request.urlopen(req, timeout=5)
                        except Exception as e:
                            return True, f"Token Valid, but failed to send to Chat ID: {e}"
                    return True, "Connection Successful!"
                return False, f"API Error: {response.getcode()}"
        except Exception as e:
            return False, str(e)
            
    def fetch_recent_chat_id(self, token):
        """Try to fetch the latest chat ID from updates"""
        try:
            url = f"{self.base_url}{token}/getUpdates"
            with urllib.request.urlopen(url, timeout=10) as response:
                if response.getcode() == 200:
                    data = json.loads(response.read().decode())
                    if data.get("result"):
                        # Look for the most recent message info
                        # We iterate backwards to find a valid chat id
                        for update in reversed(data["result"]):
                            if "message" in update:
                                return str(update["message"]["chat"]["id"])
                            elif "my_chat_member" in update:
                                return str(update["my_chat_member"]["chat"]["id"])
            return None
        except Exception as e:
            print(f"Fetch ID error: {e}")
            return None
