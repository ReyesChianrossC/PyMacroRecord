from os import path, system
from sys import platform
from utils.get_file import resource_path

try:
    from win10toast import ToastNotifier
except Exception:
    print("Not on windows. win10toast not imported.")


def show_toast(msg, title="PyMacroRecord", duration=3):
    """Show a generic toast notification"""
    if platform == "win32":
        try:
            from win10toast import ToastNotifier
            toast = ToastNotifier()
            # Note: show_toast can be blocking. Check if threaded=True is needed.
            # win10toast supports threaded=True
            toast.show_toast(
                title=title,
                msg=msg,
                duration=duration,
                icon_path=resource_path(path.join("assets", "logo.ico")),
                threaded=True
            )
        except:
            pass

    elif "linux" in platform.lower():
        system(f"""notify-send -u normal "{title}" "{msg}" """)
    elif "darwin" in platform.lower():
        system(f"""osascript -e 'display notification "{msg}" with title "{title}"'""")


def show_notification_minim(main_app):
    """Old function call for minimization notification"""
    msg = main_app.text_content["options_menu"]["settings_menu"]["minimization_toast"]
    show_toast(msg)
