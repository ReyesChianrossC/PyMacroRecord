import sys
import os
import tempfile
from os import path

# Master Lock: Prevent double launching
LOCK_FILE = path.join(tempfile.gettempdir(), "pymacrorecord_active.lock")

def check_single_instance():
    if path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r") as f:
                old_pid = int(f.read().strip())
            
            # Check if process is running using os.kill (0 signal)
            try:
                os.kill(old_pid, 0)
                # If no error, process exists
                print(f"DEBUG: PyMacroRecord is already running (PID: {old_pid}). Exiting.")
                sys.exit(0)
            except OSError as e:
                # Access Denied (WinError 5) means process exists but we can't touch it -> It runs!
                if getattr(e, 'winerror', None) == 5 or e.errno == 13: # 13 is EACCES
                    print(f"DEBUG: PyMacroRecord is running (Access Denied to PID: {old_pid}). Exiting.")
                    sys.exit(0)
                # ProcessLookupError usually means it's gone
                pass
        except Exception:
            pass # Ignore lock errors and proceed as fallback
    
    # Create/Overwrite lock
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))

if __name__ == "__main__":
    import sys
    try:
        with open("launch_log.txt", "a") as f:
            f.write(f"Launch: PID={os.getpid()} ARGS={sys.argv}\n")
    except: pass
    
    from multiprocessing import freeze_support
    freeze_support()
    
    check_single_instance()
    
    from sys import platform
    if platform.lower() == "win32":
        import ctypes
        PROCESS_PER_MONITOR_DPI_AWARE = 2
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
        except:
            pass

    from windows import MainApp
    MainApp()