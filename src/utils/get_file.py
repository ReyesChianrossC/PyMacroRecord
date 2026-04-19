import sys
from os import path

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        # Check if we are running from source and 'src' exists
        if path.exists(path.join(path.abspath("."), "src")):
             base_path = path.join(path.abspath("."), "src")
        else:
             base_path = path.abspath(".")
    return path.join(base_path, relative_path)
