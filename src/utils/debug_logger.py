import os
from datetime import datetime

class DebugLogger:
    """Simple file-based debug logger for tracking macro execution"""
    
    def __init__(self, log_dir=None):
        if log_dir is None:
            # Default to user's desktop for easy access
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            log_dir = os.path.join(desktop, "PYMACRO w SOUND")
        
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        self.log_file = os.path.join(log_dir, "pymacro_debug.log")
        
        # Clear old log on initialization
        with open(self.log_file, 'w') as f:
            f.write(f"=== PyMacro Debug Log Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n\n")
    
    def log(self, message, level="INFO"):
        """Write a message to the log file"""
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]  # Include milliseconds
        log_entry = f"[{timestamp}] [{level}] {message}\n"
        
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(log_entry)
        except Exception as e:
            print(f"Failed to write to log: {e}")
    
    def separator(self):
        """Add a visual separator"""
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write("-" * 80 + "\n")

# Create global instance
debug_logger = DebugLogger()
