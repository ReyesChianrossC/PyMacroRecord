import json
import sys
import tempfile
from tkinter import *
import tkinter as tk

from utils.not_windows import NotWindows
from windows.window import Window
from windows.main.menu_bar import MenuBar
from utils.user_settings import UserSettings
from utils.get_file import resource_path
from utils.warning_pop_up_save import confirm_save
from utils.record_file_management import RecordFileManagement
from utils.version import Version
from windows.others.new_ver_avalaible import NewVerAvailable
from hotkeys.hotkeys_manager import HotkeysManager
from macro import Macro
from utils.debug_logger import debug_logger
from utils.telegram_notifier import TelegramNotifier
from os import path
from sys import platform, argv
from pystray import Icon
from pystray import MenuItem
from PIL import Image
from threading import Thread
from json import load
from time import time
import copy
import os
import glob
import random

if platform.lower() == "win32":
    from tkinter.ttk import *
from tkinter import filedialog, messagebox, scrolledtext
from windows.main.area_selector import AreaSelector, MultiAreaSelector
from datetime import datetime

def deepcopy_dict_missing_entries(dst:dict,src:dict):
# recursively copy entries that are in src but not in dst
    for k,v in src.items():
        if k not in dst:
            dst[k] = copy.deepcopy(v)
        elif isinstance(v,dict):
            deepcopy_dict_missing_entries(dst[k],v)

class ScriptListbox(Frame):
    def __init__(self, parent, main_app, initial_path=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.main_app = main_app
        # Use relative path for portability or custom path
        if initial_path:
             self.scripts_path = initial_path
        else:
             self.scripts_path = path.abspath(path.join(path.dirname(__file__), "..", "..", "scripts"))
        
        # Create the main frame with title
        self.title_label = Label(self, text="Recorded Scripts", font=("Arial", 10, "bold"))
        self.title_label.pack(pady=(5, 5))
        
        # Create frame for listbox and scrollbar
        self.listbox_frame = Frame(self)
        self.listbox_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        # Create listbox with scrollbar
        self.scrollbar = Scrollbar(self.listbox_frame)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        
        self.listbox = Listbox(
            self.listbox_frame,
            yscrollcommand=self.scrollbar.set,
            selectmode=SINGLE,
            height=4,
            font=("Arial", 9)
        )
        self.listbox.pack(side=LEFT, fill=BOTH, expand=True)
        self.scrollbar.config(command=self.listbox.yview)
        
        # Bind double-click event
        self.listbox.bind("<Double-Button-1>", self.on_script_double_click)
        
        # Create buttons frame
        self.buttons_frame = Frame(self)
        self.buttons_frame.pack(fill=X, padx=5, pady=2)
        
        # Refresh button
        self.refresh_btn = Button(
            self.buttons_frame,
            text="Refresh",
            command=self.refresh_script_list,
            width=8
        )
        self.refresh_btn.pack(side=LEFT, padx=(0, 2))
        
        # Play button
        self.play_btn = Button(
            self.buttons_frame,
            text="Play",
            command=self.play_selected_script,
            width=8
        )
        self.play_btn.pack(side=LEFT, padx=2)
        
        # Delete button
        self.delete_btn = Button(
            self.buttons_frame,
            text="Delete",
            command=self.delete_selected_script,
            width=8
        )
        self.delete_btn.pack(side=RIGHT)
        
        # Status label
        self.status_label = Label(self, text="Ready", font=("Arial", 8))
        self.status_label.pack(pady=(2, 0))
        
        # Load scripts on initialization
        self.refresh_script_list()
    
    def refresh_script_list(self):
        """Refresh the list of available scripts"""
        try:
            # Clear current listbox
            self.listbox.delete(0, END)
            
            # Create scripts directory if it doesn't exist
            if not os.path.exists(self.scripts_path):
                os.makedirs(self.scripts_path)
                self.status_label.config(text="Scripts folder created")
                return
            
            # Get all script files (looking for .pmr files only)
            script_extensions = ['*.pmr']
            script_files = []
            
            for extension in script_extensions:
                all_found = glob.glob(os.path.join(self.scripts_path, extension))
                # Filter out sp1.pmr and sp2.pmr - UNHIDDEN per request
                # This allow user to see and verify special case macros in the list
                script_files.extend(all_found)
            
            if not script_files:
                self.listbox.insert(END, "No scripts found")
                self.status_label.config(text="No scripts found in folder")
                return
            
            # Add script names to listbox (without full path and extension for cleaner look)
            for script_file in sorted(script_files):
                script_name = os.path.splitext(os.path.basename(script_file))[0]
                display_name = f"{script_name} ({os.path.splitext(script_file)[1]})"
                self.listbox.insert(END, display_name)
            
            self.status_label.config(text=f"Found {len(script_files)} scripts")
            
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
    
    def on_script_double_click(self, event):
        """Handle double-click on script"""
        self.play_selected_script()
    
    def play_selected_script(self):
        """Play the selected script"""
        selection = self.listbox.curselection()
        if not selection:
            self.status_label.config(text="No script selected")
            return
        
        script_display_name = self.listbox.get(selection[0])
        if script_display_name == "No scripts found":
            return
        
        # Extract actual filename from display name
        script_name = script_display_name.split(" (")[0]
        
        # Find the actual file with extension
        script_files = []
        for ext in ['*.pmr']:
            script_files.extend(glob.glob(os.path.join(self.scripts_path, ext)))
        
        actual_file = None
        for file_path in script_files:
            if os.path.splitext(os.path.basename(file_path))[0] == script_name:
                actual_file = file_path
                break
        
        if not actual_file:
            self.status_label.config(text="Script file not found")
            return
        
        try:
            self.status_label.config(text=f"Playing: {script_name}")
            
            # Load and play the macro
            if actual_file.endswith('.pmr') or actual_file.endswith('.json'):
                with open(actual_file, 'r') as record:
                    loaded_content = load(record)
                
                # Import the record into the main app's macro
                self.main_app.macro.import_record(loaded_content)
                self.main_app.macro_recorded = True
                self.main_app.macro_saved = True
                self.main_app.playBtn.config(state=NORMAL)
                
                # Start playback immediately
                if self.main_app.loop_scripts_var.get():
                    # Determine loop limit for the current script
                    loop_limit = self.main_app.settings.settings_dict["Playback"].get("Loop_Scripts_Limit", 0)
                    if self.main_app.loops_done < loop_limit or loop_limit == 0:
                        self.main_app.log(f"Loop {self.main_app.loops_done}/{loop_limit if loop_limit > 0 else '∞'}", "info")
                        # Automated resumption - pass manual_start=False
                        self.main_app.after(50, lambda: self.main_app.on_play_click(manual_start=False))
                        return
                else:
                    self.main_app.macro.start_playback()
                
            else:
                self.status_label.config(text="Unsupported file format")
            
        except Exception as e:
            self.status_label.config(text=f"Error playing script: {str(e)}")
    
    def delete_selected_script(self):
        """Delete the selected script file"""
        selection = self.listbox.curselection()
        if not selection:
            self.status_label.config(text="No script selected")
            return
        
        script_display_name = self.listbox.get(selection[0])
        if script_display_name == "No scripts found":
            return
        
        script_name = script_display_name.split(" (")[0]
        
        # Find the actual file
        script_files = []
        for ext in ['*.pmr']:
            script_files.extend(glob.glob(os.path.join(self.scripts_path, ext)))
        
        actual_file = None
        for file_path in script_files:
            if os.path.splitext(os.path.basename(file_path))[0] == script_name:
                actual_file = file_path
                break
        
        if not actual_file:
            self.status_label.config(text="Script file not found")
            return
        
        # Ask for confirmation
        from tkinter import messagebox
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{script_name}'?"):
            try:
                os.remove(actual_file)
                self.refresh_script_list()
                self.status_label.config(text=f"Deleted: {script_name}")
            except Exception as e:
                self.status_label.config(text=f"Error deleting script: {str(e)}")


class MainApp(Window):
    """Main windows of the application"""

    def __init__(self):
        # Manageable window size for most screens
        super().__init__("PyMacroRecord", 550, 650)
        self.resizable(True, True)
        self.attributes("-topmost", 1)
        if platform == "win32":
            self.iconbitmap(resource_path(path.join("assets", "logo.ico")))

        self.settings = UserSettings(self)
        
        # Check for Last Script Folder
        last_folder = self.settings.settings_dict["Others"].get("Last_Script_Folder", "")
        self.initial_scripts_path = last_folder if (last_folder and os.path.exists(last_folder)) else None

        self.load_language()

        # For save message purpose
        self.macro_saved = False
        self.macro_recorded = False
        self.current_file = None
        self.prevent_record = False

        # Clear any stale quit signal left over from a crash
        try:
            quit_signal = os.path.join(tempfile.gettempdir(), "pymacro_quit.signal")
            if os.path.exists(quit_signal):
                os.remove(quit_signal)
        except Exception:
            pass

        # Loop state variable
        # Initialize from settings
        is_infinite = self.settings.settings_dict["Playback"]["Repeat"]["Infinite"]
        self.loop_var = BooleanVar(value=is_infinite)
        
        is_loop_scripts = self.settings.settings_dict["Playback"].get("Loop_Scripts", False)
        self.loop_scripts_var = BooleanVar(value=is_loop_scripts)
        
        # Global loop interval (in milliseconds)
        self.global_loop_interval = self.settings.settings_dict["Playback"].get("Global_Loop_Interval", 0)

        # Playlist variables
        self.playlist = []
        self.current_playlist_index = 0
        self.is_playlist_playing = False
        self.total_playlist_runs = 0
        self.loops_done = 0
        self.overlay = None

        # Case-specific state variables
        self.case_vars = {}
        for case in ["N", "SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
            case_settings = self.settings.settings_dict["Special_Cases"][case]
            self.case_vars[case] = {
                "enabled": BooleanVar(value=case_settings.get("Enabled", False)),
                "alarm": BooleanVar(value=case_settings.get("Alarm", False)),
                "confidence": case_settings.get("Confidence", 0.75),
                "limit_val": StringVar(value=str(case_settings.get("Loop_Limit", "")) if case_settings.get("Loop_Limit") else ""),
                "stop_program": BooleanVar(value=case_settings.get("Stop_Program", False)),
                "tg_alert": BooleanVar(value=case_settings.get("TG_Alert", False)),
                "tg_message": StringVar(value=case_settings.get("TG_Message", ""))
            }

        self.image_recognition_var = self.case_vars["N"]["enabled"] # Backward compatibility for some methods
        self.alarm_var = self.case_vars["N"]["alarm"]

        self.version = Version(self.settings.settings_dict, self)

        self.menu = MenuBar(self)  # Menu Bar
        self.macro = Macro(self)

        self.validate_cmd = self.register(self.validate_input)

        self.hotkeyManager = HotkeysManager(self)
        self.telegram_notifier = TelegramNotifier(self)

        self.status_text = Label(self, text='', relief=SUNKEN, anchor=W)
        if self.settings.settings_dict["Recordings"]["Show_Events_On_Status_Bar"]:
            self.status_text.pack(side=BOTTOM, fill=X)

        # Main container with a scrollbar for vertical space
        self.main_canvas = Canvas(self, highlightthickness=0)
        self.main_scrollbar = Scrollbar(self, orient=VERTICAL, command=self.main_canvas.yview)
        self.main_container = Frame(self.main_canvas)
        
        # Configure canvas
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        
        # Packing for scroll functionality
        self.main_scrollbar.pack(side=RIGHT, fill=Y)
        self.main_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        
        # Create a window in the canvas for our frame
        self.canvas_window = self.main_canvas.create_window((0, 0), window=self.main_container, anchor=NW)
        
        # Update scrollregion when frame size changes
        self.main_container.bind("<Configure>", lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        
        # Make canvas width match parent width
        self.main_canvas.bind("<Configure>", self._on_canvas_configure)

        # Top frame for main buttons
        self.buttons_frame = Frame(self.main_container)
        self.buttons_frame.pack(fill=X, pady=(0, 5))

        # Main Buttons (Start record, stop record, start playback, stop playback)
        # Play Button
        self.playImg = PhotoImage(file=resource_path(path.join("assets", "button", "play.png")))

        # Import record if opened with .pmr extension
        if len(argv) > 1:
            with open(sys.argv[1], 'r') as record:
                loaded_content = load(record)
            self.macro.import_record(loaded_content)
            self.macro.import_record(loaded_content)
            self.playBtn = Button(self.buttons_frame, image=self.playImg, command=self.on_play_click)
            self.macro_recorded = True
            self.macro_saved = True
        else:
            self.playBtn = Button(self.buttons_frame, image=self.playImg, state=DISABLED)
        self.playBtn.pack(side=LEFT, padx=50)

        # Record Button
        self.recordImg = PhotoImage(file=resource_path(path.join("assets", "button", "record.png")))
        self.recordBtn = Button(self.buttons_frame, image=self.recordImg, command=self.macro.start_record)
        self.recordBtn.pack(side=RIGHT, padx=50)

        # Loop Checkbox
        self.loop_checkbox = Checkbutton(
            self.buttons_frame, 
            text="Loop", 
            variable=self.loop_var,
            command=self.toggle_infinite_loop
        )
        self.loop_checkbox.pack(side=LEFT, padx=10)

        # Random Loop Checkbox
        self.random_loop_checkbox = Checkbutton(
            self.buttons_frame,
            text="Random Loop",
            variable=self.loop_scripts_var,
            command=self.toggle_random_loop
        )
        self.random_loop_checkbox.pack(side=LEFT, padx=10)

        # --- Case Control Blocks (N, SP1, SP2) ---
        self.case_container = Frame(self.main_container)
        self.case_container.pack(fill=X, pady=5)
        
        for case in ["N", "SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
            self.create_case_block(self.case_container, case)

        # Alarm Volume Slider (Global)
        self.volume_frame = Frame(self.main_container)
        self.volume_frame.pack(fill=X, pady=5)
        
        # Telegram Controls (Left of Volume)
        self.tg_status_label = tk.Label(self.volume_frame, text="X", fg="red", width=2)
        self.tg_status_label.pack(side=LEFT, padx=(5, 0))
        
        # Check if already configured to set initial state
        if self.telegram_notifier.is_enabled():
            self.tg_status_label.config(text="OK", fg="green")

        Button(self.volume_frame, text="TG", width=3, command=self.open_telegram_settings).pack(side=LEFT, padx=(0, 10))
        
        Label(self.volume_frame, text="Alarm Volume:").pack(side=LEFT)
        current_volume = self.settings.settings_dict["Image_Recognition"].get("Volume", 50)
        self.volume_scale = Scale(
            self.volume_frame, 
            from_=0, 
            to=100, 
            orient=HORIZONTAL, 
            length=200,
            command=self.on_volume_change
        )
        self.volume_scale.set(current_volume)
        self.volume_scale.pack(side=LEFT, padx=10)

        # Stop Button
        self.stopImg = PhotoImage(file=resource_path(path.join("assets", "button", "stop.png")))

        # Add the script listbox below the main buttons
        self.script_listbox = ScriptListbox(self.main_container, self, initial_path=self.initial_scripts_path)
        self.script_listbox.pack(fill=BOTH, expand=True)
        
        # Global Loop Interval controls
        self.global_interval_frame = Frame(self.main_container)
        self.global_interval_frame.pack(fill=X, pady=5, padx=5)
        
        Label(self.global_interval_frame, text="Global Loop Interval (ms):", width=22, anchor=W).pack(side=LEFT)
        self.global_interval_var = StringVar(value=str(self.global_loop_interval) if self.global_loop_interval > 0 else "")
        Entry(self.global_interval_frame, textvariable=self.global_interval_var, width=10,
              validate="key", validatecommand=self.validate_cmd).pack(side=LEFT, padx=2)
        
        Button(self.global_interval_frame, text="Set", width=6, command=self.set_global_interval).pack(side=LEFT, padx=2)
        Button(self.global_interval_frame, text="Remove", width=6, command=self.remove_global_interval).pack(side=LEFT, padx=2)

        # Load batch settings for the initial folder
        try:
            self.load_batch_settings(self.script_listbox.scripts_path)
        except Exception as e:
            print(f"Startup detected error loading batch settings: {e}")

        record_management = RecordFileManagement(self, self.menu)

        self.bind('<Control-Shift-S>', record_management.save_macro_as)
        self.bind('<Control-s>', record_management.save_macro)
        self.bind('<Control-l>', record_management.load_macro)
        self.bind('<Control-n>', record_management.new_macro)

        self.protocol("WM_DELETE_WINDOW", self.quit_software)
        if platform.lower() != "darwin":
            Thread(target=self.systemTray).start()

        self.attributes("-topmost", 0)

        if platform != "win32" and self.settings.first_time:
            NotWindows(self)

        if self.settings.settings_dict["Others"]["Check_update"]:
            if self.version.new_version != "" and self.version.version != self.version.new_version:
                if time() > self.settings.settings_dict["Others"]["Remind_new_ver_at"]:
                    NewVerAvailable(self, self.version.new_version)
        # Start polling for /reset signal from tg_remote.py
        self._start_reset_signal_polling()

        try:
            self.mainloop()
        except Exception as e:
            messagebox.showerror("Critical Error", f"Application crashed:\n{e}")
            raise e

    # ── Telegram Remote: signal polling ──────────────────────────────────────
    _SIGNAL_RESET    = os.path.join(tempfile.gettempdir(), "pymacro_reset.signal")
    _SIGNAL_STOP     = os.path.join(tempfile.gettempdir(), "pymacro_stop.signal")
    _SIGNAL_SETLIMIT = os.path.join(tempfile.gettempdir(), "pymacro_setlimit.signal")
    _SIGNAL_START    = os.path.join(tempfile.gettempdir(), "pymacro_start.signal")

    def _start_reset_signal_polling(self):
        """Begin polling every 2 seconds for signals from tg_remote.py"""
        self._poll_reset_signal()

    def _poll_reset_signal(self):
        """Check for all tg_remote signal files and act on them, then reschedule."""
        try:
            # ── /reset: stop + restart ────────────────────────────────────────
            if os.path.exists(self._SIGNAL_RESET):
                os.remove(self._SIGNAL_RESET)
                print("[tg_remote] /reset signal — restarting playback")
                if self.macro.playback:
                    self.macro.stop_playback(True)
                    self.after(800, lambda: self.on_play_click(manual_start=True))
                else:
                    self.on_play_click(manual_start=True)

            # ── /stop: stop playback ──────────────────────────────────────────
            if os.path.exists(self._SIGNAL_STOP):
                os.remove(self._SIGNAL_STOP)
                print("[tg_remote] /stop signal — stopping playback")
                if self.macro.playback:
                    self.macro.stop_playback(True)

            # ── /setLimit: update the Case N loop limit ───────────────────────
            if os.path.exists(self._SIGNAL_SETLIMIT):
                try:
                    with open(self._SIGNAL_SETLIMIT, "r") as f:
                        raw = f.read().strip()
                    os.remove(self._SIGNAL_SETLIMIT)
                    limit = int(raw)
                    if limit > 0:
                        # Stop current playback first if running
                        if self.macro.playback:
                            self.macro.stop_playback(True)
                        # Remove old limit, set new one
                        self.settings.change_settings("Special_Cases", "N", "Loop_Limit", limit)
                        # Sync to UI if the limit_val widget exists
                        if hasattr(self, "case_vars") and "N" in self.case_vars:
                            self.case_vars["N"]["limit_val"].set(str(limit))
                        print(f"[tg_remote] /setLimit — loop limit set to {limit}")
                except Exception as e:
                    print(f"[tg_remote] setlimit signal error: {e}")

            # ── /start: start playback (after setLimit confirmation) ──────────
            if os.path.exists(self._SIGNAL_START):
                os.remove(self._SIGNAL_START)
                print("[tg_remote] /start signal — starting playback")
                if not self.macro.playback:
                    self.on_play_click(manual_start=True)

        except Exception as e:
            print(f"[tg_remote] Poll error: {e}")

        # Reschedule every 2000ms
        self.after(2000, self._poll_reset_signal)
    # ─────────────────────────────────────────────────────────────────────────


    def log(self, message, tag="info"):
        """No-op log method (Activity Log removed)"""
        pass

    def load_language(self) -> None:
        self.lang = self.settings.settings_dict["Language"]
        with open(resource_path(path.join('langs', self.lang + '.json')), encoding='utf-8') as f:
            self.text_content = json.load(f)
        self.text_content = self.text_content["content"]

        if self.lang != "en":
            with open(resource_path(path.join('langs', 'en.json')), encoding='utf-8') as f:
                en = json.load(f)
            deepcopy_dict_missing_entries(self.text_content, en["content"])

    def systemTray(self):
        """Just to show little icon on system tray"""
        image = Image.open(resource_path(path.join("assets", "logo.ico")))
        menu = (
            MenuItem('Show', action=self.deiconify, default=True),
        )
        self.icon = Icon("name", image, "PyMacroRecord", menu)
        self.icon.run()

    def validate_input(self, action, value_if_allowed):
        """Prevents from adding letters on an Entry label"""
        if action == "1":  # Insert
            try:
                float(value_if_allowed)
                return True
            except ValueError:
                return False
        return True

    def quit_software(self, force=False):
        if not self.macro_saved and self.macro_recorded and not force:
            wantToSave = confirm_save(self)
            if wantToSave:
                RecordFileManagement(self, self.menu).save_macro()
            elif wantToSave == None:
                return
        if platform.lower() != "darwin":
            self.icon.stop()
        if platform.lower() == "linux":
            self.destroy()

        # Tell tg_remote.py to shut down cleanly
        try:
            quit_signal = os.path.join(tempfile.gettempdir(), "pymacro_quit.signal")
            with open(quit_signal, "w") as f:
                f.write("quit")
        except Exception as e:
            print(f"Failed to write quit signal for tg_remote: {e}")

        self.quit()

    def toggle_infinite_loop(self):
        """Toggle the infinite loop setting when checkbox is clicked"""
        is_infinite = self.loop_var.get()
        # Update settings using the change_settings method
        self.settings.change_settings("Playback", "Repeat", "Infinite", is_infinite)


    def sync_ui_with_settings(self):
        """Load settings from JSON into the UI variables"""
        for case in ["N", "SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
            c_sett = self.settings.settings_dict.get("Special_Cases", {}).get(case, {})
            self.case_vars[case]["enabled"].set(c_sett.get("Enabled", False))
            self.case_vars[case]["alarm"].set(c_sett.get("Alarm", False))
            if "stop_program" in self.case_vars[case]:
                self.case_vars[case]["stop_program"].set(c_sett.get("Stop_Program", False))
            if "tg_alert" in self.case_vars[case]:
                self.case_vars[case]["tg_alert"].set(c_sett.get("TG_Alert", False))
            if "tg_message" in self.case_vars[case]:
                self.case_vars[case]["tg_message"].set(c_sett.get("TG_Message", ""))
            self.case_vars[case]["confidence"] = c_sett.get("Confidence", 0.75)
            
            limit = c_sett.get("Loop_Limit")
            if limit and limit is not False:
                self.case_vars[case]["limit_val"].set(str(limit))
            else:
                self.case_vars[case]["limit_val"].set("")
            
            # Update label if it exists
            if "conf_label" in self.case_vars[case]:
                self.case_vars[case]["conf_label"].configure(text=f"{int(self.case_vars[case]['confidence'] * 100)}%")

    def toggle_image_recognition(self, case="N"):
        """Toggle the image recognition setting"""
        enabled = self.case_vars[case]["enabled"].get()
        self.settings.change_settings("Special_Cases", case, "Enabled", enabled)
        
        if enabled:
            # Check if image is set
            image_path = self.settings.settings_dict["Special_Cases"][case]["Image_Path"]
            if not image_path or not path.exists(image_path):
                if messagebox.askyesno("Image Not Found", f"No image selected for {case}. Do you want to upload one now?"):
                    self.upload_image(case)
                else:
                    self.case_vars[case]["enabled"].set(False)
                    self.settings.change_settings("Special_Cases", case, "Enabled", False)
                    return

            # Check if area is set (now a list of areas)
            areas = self.settings.settings_dict["Special_Cases"][case]["Area"]
            if not areas or not isinstance(areas, list) or len(areas) == 0:
                if messagebox.askyesno("Area Not Set", f"No screen area defined for {case}. Do you want to select the area now?"):
                    self.select_area(case)
                else:
                    self.case_vars[case]["enabled"].set(False)
                    self.settings.change_settings("Special_Cases", case, "Enabled", False)
                    return

    def toggle_stop_program(self, case):
        """Toggle the stop program setting for SP1/SP2"""
        enabled = self.case_vars[case]["stop_program"].get()
        self.settings.change_settings("Special_Cases", case, "Stop_Program", enabled)

    def toggle_alarm(self, case="N"):
        """Toggle the alarm setting"""
        enabled = self.case_vars[case]["alarm"].get()
        self.settings.change_settings("Special_Cases", case, "Alarm", enabled)

    def toggle_tg_alert(self, case):
        """Toggle the TG Alert setting"""
        enabled = self.case_vars[case]["tg_alert"].get()
        self.settings.change_settings("Special_Cases", case, "TG_Alert", enabled)

    def set_tg_message(self, case):
        """Set custom Telegram message"""
        msg = self.case_vars[case]["tg_message"].get().strip()
        self.settings.change_settings("Special_Cases", case, "TG_Message", msg)
        messagebox.showinfo("Message Set", f"Custom message for {case} saved.")

    def remove_tg_message(self, case):
        """Remove custom Telegram message"""
        self.case_vars[case]["tg_message"].set("")
        self.settings.change_settings("Special_Cases", case, "TG_Message", "")
        messagebox.showinfo("Message Removed", f"Custom message for {case} removed. Using default.")

    def on_volume_change(self, value):
        """Handle volume change"""
        self.settings.change_settings("Image_Recognition", "Volume", None, int(float(value)))

    def upload_image(self, case="N"):
        """Handle image upload for a specific case"""
        file_path = filedialog.askopenfilename(
            title=f"Select Image for {case}",
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp")]
        )
        
        if file_path:
            current_folder = self.script_listbox.scripts_path
            if not os.path.exists(current_folder):
                os.makedirs(current_folder)
            
            # Determine target filename
            filename_map = {
                "N": "target.png",
                "SP1": "sp1target.png",
                "SP2": "sp2target.png",
                "SP3": "sp3target.png",
                "SP4": "sp4target.png",
                "SP5": "sp5target.png",
                "SP6": "sp6target.png",
            }
            filename = filename_map.get(case, f"{case.lower()}target.png")
            target_path = path.join(current_folder, filename)
            
            try:
                img = Image.open(file_path)
                img.save(target_path)
                
                self.settings.change_settings("Special_Cases", case, "Image_Path", target_path)
                messagebox.showinfo("Success", f"Image saved for {case}:\n{target_path}")
                
                if self.case_vars[case]["enabled"].get():
                   self.toggle_image_recognition(case) 
                   
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save image: {e}")

    def select_area(self, case="N"):
        """Open multi-area selector"""
        self.withdraw()
        MultiAreaSelector(self, lambda areas: self.on_area_selected(areas, case))
    
    def on_area_selected(self, areas, case="N"):
        """Callback for area selection — areas is a list of [x1,y1,x2,y2] lists"""
        self.deiconify()
        if areas:  # non-empty list means user confirmed
            self.settings.change_settings("Special_Cases", case, "Area", areas)
            self.save_batch_settings(case)
            self.after(100, lambda: messagebox.showinfo(
                "Area Selected",
                f"{len(areas)} area(s) set for {case} monitoring."
            ))
        # If areas == [] the user cancelled — do nothing

    def toggle_random_loop(self):
        """Toggle random loop scripts"""
        is_loop = self.loop_scripts_var.get()
        self.settings.change_settings("Playback", "Loop_Scripts", None, is_loop)
        
        if is_loop or self.macro_recorded:
            self.playBtn.config(state=NORMAL)
        else:
            self.playBtn.config(state=DISABLED)

    def set_loop_limit(self, case="N"):
        """Set the hard loop limit for a case"""
        val = self.case_vars[case]["limit_val"].get()
        if not val:
            messagebox.showwarning("Invalid Input", f"Please enter an integer for the {case} loop limit.")
            return
        
        try:
            limit = int(val)
            if limit <= 0:
                messagebox.showwarning("Invalid Input", "Limit must be greater than 0.")
                return
            
            self.settings.change_settings("Special_Cases", case, "Loop_Limit", limit)
            if self.macro.playback:
                self.macro.stop_playback(True)
            
            messagebox.showinfo("Limit Set", f"{case} loop limit set to {limit}. Ongoing playback halted.")
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid integer.")

    def remove_loop_limit(self, case="N"):
        """Remove the loop limit for a case"""
        self.settings.change_settings("Special_Cases", case, "Loop_Limit", None)
        self.case_vars[case]["limit_val"].set("")
        messagebox.showinfo("Limit Removed", f"{case} loop limit removed.")

    def adjust_confidence(self, delta, case="N"):
        """Increase or decrease confidence for a case"""
        current = self.settings.settings_dict["Special_Cases"][case].get("Confidence", 0.75)
        new_val = round(max(0.60, min(1.0, current + delta)), 2)
        self.settings.change_settings("Special_Cases", case, "Confidence", new_val)
        self.case_vars[case]["confidence"] = new_val
        if "conf_label" in self.case_vars[case]:
            self.case_vars[case]["conf_label"].config(text=f"{int(new_val * 100)}%")
    
    def set_global_interval(self):
        """Set the global loop interval"""
        val = self.global_interval_var.get()
        if not val:
            messagebox.showwarning("Invalid Input", "Please enter an integer for the global loop interval (in milliseconds).")
            return
        
        try:
            interval = int(val)
            if interval < 0:
                messagebox.showwarning("Invalid Input", "Interval must be 0 or greater.")
                return
            
            self.global_loop_interval = interval
            self.settings.change_settings("Playback", "Global_Loop_Interval", None, interval)
            messagebox.showinfo("Interval Set", f"Global loop interval set to {interval}ms.")
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid integer.")
    
    def remove_global_interval(self):
        """Remove the global loop interval"""
        self.global_loop_interval = 0
        self.settings.change_settings("Playback", "Global_Loop_Interval", None, 0)
        self.global_interval_var.set("")
        messagebox.showinfo("Interval Removed", "Global loop interval removed.")

    def create_case_block(self, parent, case):
        """Create a block of controls for a specific case (N, SP1, SP2)"""
        frame = LabelFrame(parent, text=f" Case {case} ")
        frame.pack(fill=X, pady=2, padx=5)

        # First row: Image Stop & Alarm
        row1 = Frame(frame)
        row1.pack(fill=X)
        
        label_text = "Image Stop" if case == "N" else "Image Trigger"
        Checkbutton(row1, text=label_text, variable=self.case_vars[case]["enabled"], 
                    command=lambda: self.toggle_image_recognition(case)).pack(side=LEFT, padx=5)
        Checkbutton(row1, text="Alarm", variable=self.case_vars[case]["alarm"], 
                    command=lambda: self.toggle_alarm(case)).pack(side=LEFT, padx=5)

        if case in ["SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
             Checkbutton(row1, text="Stop Program", variable=self.case_vars[case]["stop_program"], 
                    command=lambda: self.toggle_stop_program(case)).pack(side=LEFT, padx=5)

        # Custom Message Controls (Right aligned)
        # We use a frame to group them
        msg_frame = Frame(row1)
        msg_frame.pack(side=RIGHT, padx=5)
        
        Checkbutton(msg_frame, text="TG Alert", variable=self.case_vars[case]["tg_alert"],
                    command=lambda: self.toggle_tg_alert(case)).pack(side=RIGHT, padx=(5, 0))
        
        tk.Button(msg_frame, text="Remove", width=6, command=lambda: self.remove_tg_message(case)).pack(side=RIGHT, padx=2)
        tk.Button(msg_frame, text="Set", width=4, command=lambda: self.set_tg_message(case)).pack(side=RIGHT, padx=2)
        Entry(msg_frame, textvariable=self.case_vars[case]["tg_message"], width=15).pack(side=RIGHT, padx=2)

        # Second row: Buttons
        row2 = Frame(frame)
        row2.pack(fill=X, pady=2)
        
        Button(row2, text="Upload Image", width=12, command=lambda: self.upload_image(case)).pack(side=LEFT, padx=2)
        Button(row2, text="Select Area", width=12, command=lambda: self.select_area(case)).pack(side=LEFT, padx=2)
        
        # Confidence Level
        Label(row2, text="Conf:").pack(side=LEFT, padx=5)
        Button(row2, text="▲", width=2, command=lambda: self.adjust_confidence(0.05, case)).pack(side=LEFT)
        Button(row2, text="▼", width=2, command=lambda: self.adjust_confidence(-0.05, case)).pack(side=LEFT)
        
        conf_label = Label(row2, text=f"{int(self.case_vars[case]['confidence'] * 100)}%")
        conf_label.pack(side=LEFT, padx=2)
        self.case_vars[case]["conf_label"] = conf_label # Store for easy update

        if case == "N":
            # Third row: Loop Limit
            row3 = Frame(frame)
            row3.pack(fill=X, pady=2)
            
            Label(row3, text="Loop Limit:", width=10, anchor=W).pack(side=LEFT)
            Entry(row3, textvariable=self.case_vars[case]["limit_val"], width=10, 
                  validate="key", validatecommand=self.validate_cmd).pack(side=LEFT, padx=2)
            
            Button(row3, text="Set", width=6, command=lambda: self.set_loop_limit(case)).pack(side=LEFT, padx=2)
            Button(row3, text="Remove", width=6, command=lambda: self.remove_loop_limit(case)).pack(side=LEFT, padx=2)

    def open_telegram_settings(self):
        """Open dialog to configure Telegram"""
        dialog = Toplevel(self)
        dialog.title("Telegram Settings")
        dialog.geometry("400x250")
        dialog.transient(self)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        Label(dialog, text="Telegram Configuration", font=("Arial", 10, "bold")).pack(pady=10)

        # Token
        f1 = Frame(dialog)
        f1.pack(fill=X, padx=10, pady=5)
        Label(f1, text="Bot Token:").pack(side=LEFT)
        token_entry = Entry(f1, width=30)
        token_entry.pack(side=RIGHT, expand=True, fill=X)
        token_entry.insert(0, self.settings.settings_dict.get("Others", {}).get("Telegram_Token", ""))

        # Chat ID
        f2 = Frame(dialog)
        f2.pack(fill=X, padx=10, pady=5)
        Label(f2, text="Chat ID:").pack(side=LEFT)
        chat_id_entry = Entry(f2, width=30)
        chat_id_entry.pack(side=RIGHT, expand=True, fill=X)
        chat_id_entry.insert(0, self.settings.settings_dict.get("Others", {}).get("Telegram_Chat_ID", ""))

        # Helper to strict clean inputs (remove invisible unicode like \u200e LRM)
        def clean_input(text):
            # Keep only ASCII non-whitespace characters
            return ''.join(c for c in text if c.isascii() and not c.isspace())

        # Helper to fetch ID
        def auto_fetch_id():
            raw_token = token_entry.get()
            token = clean_input(raw_token)
            
            # Update entry if cleaned
            if token != raw_token:
                token_entry.delete(0, END)
                token_entry.insert(0, token)

            if not token:
                messagebox.showerror("Error", "Please enter Bot Token first.", parent=dialog)
                return
            
            messagebox.showinfo("Instructions", "Please send a message (e.g., 'Hello') to your bot on Telegram NOW.\n\nClick OK after you have sent it.", parent=dialog)
            
            chat_id = self.telegram_notifier.fetch_recent_chat_id(token)
            if chat_id:
                chat_id_entry.delete(0, END)
                chat_id_entry.insert(0, chat_id)
                messagebox.showinfo("Success", f"Found Chat ID: {chat_id}", parent=dialog)
            else:
                messagebox.showerror("Failed", "Could not find any recent messages. Make sure you messaged the correct bot.", parent=dialog)

        tk.Button(dialog, text="Auto-Detect Chat ID", command=auto_fetch_id).pack(pady=5)

        # Test & Save
        def save_and_test():
            raw_token = token_entry.get()
            raw_chat_id = chat_id_entry.get()
            
            token = clean_input(raw_token)
            chat_id = clean_input(raw_chat_id)
            
            # Update entries if cleaned
            if token != raw_token:
                token_entry.delete(0, END)
                token_entry.insert(0, token)
            if chat_id != raw_chat_id:
                chat_id_entry.delete(0, END)
                chat_id_entry.insert(0, chat_id)
            
            success, msg = self.telegram_notifier.test_connection(token, chat_id)
            if success:
                self.settings.change_settings("Others", "Telegram_Token", None, token)
                self.settings.change_settings("Others", "Telegram_Chat_ID", None, chat_id)
                self.update_tg_status_icon(True)
                
                # Save to batch config immediately as per user request
                self.save_batch_settings()
                
                messagebox.showinfo("Success", msg, parent=dialog)
                dialog.destroy()
            else:
                self.update_tg_status_icon(False)
                messagebox.showerror("Connection Failed", msg, parent=dialog)

        tk.Button(dialog, text="Test & Save", command=save_and_test, width=15, bg="#dddddd").pack(pady=15)

    def update_tg_status_icon(self, is_connected):
        """Update the visual indicator for Telegram status"""
        if hasattr(self, 'tg_status_label'):
            if is_connected:
                self.tg_status_label.config(text="OK", fg="green")
            else:
                self.tg_status_label.config(text="X", fg="red")

    def send_telegram_alert(self, case):
        """Helper to send alerts from macro.py"""
        # Check if BOTH the global system is enabled AND the specific case checkbox is ticked
        if self.telegram_notifier.is_enabled() and self.case_vars[case]["tg_alert"].get():
            custom_msg = self.case_vars[case]["tg_message"].get().strip()
            if custom_msg:
                 self.telegram_notifier.send_message(custom_msg)
            else:
                 self.telegram_notifier.send_message(f"[PyMacro] Alert: Target Detected for Case {case}!")

    def on_play_click(self, manual_start=True):
        """Handle play button click (manual or automated)"""
        debug_logger.log(f"ON_PLAY_CLICK: manual_start={manual_start}, loops_done={self.loops_done}")
        if self.macro.playback or self.macro.record:
            return
        
        # Only reset the hard stop flag if the user MANUALLY clicked Play
        if manual_start:
            self.macro.hard_stop_triggered = False
            # Only reset counter on manual start, not automated resumptions
            debug_logger.log(f"Manual start detected - RESETTING loops_done from {self.loops_done} to 0")
            self.loops_done = 0
            self.total_playlist_runs = 0
        else:
            debug_logger.log(f"Automated resumption - PRESERVING loops_done={self.loops_done}")
        
        # Absolute guard: If we are in a hard stop state and this is NOT a manual start, ABORT
        if self.macro.hard_stop_triggered and not manual_start:
            self.log("Automated play blocked (Hard Stop active)", "stop")
            return

        self.macro.active_case = "N" # Default case
        debug_logger.log(f"About to start playback: loops_done={self.loops_done}")
        if self.loop_scripts_var.get():
             self.start_playlist_playback()
        else:
             self.macro.start_playback()
        debug_logger.log(f"After playback started: loops_done={self.loops_done}")

    def on_playback_finished(self):
        """Called when a macro finishes playing naturally"""
        debug_logger.log(f"on_playback_finished called. Active Case: {getattr(self.macro, 'active_case', 'N')}")
        if getattr(self.macro, "hard_stop_triggered", False):
            debug_logger.log("on_playback_finished: Hard stop active, stopping playlist.")
            self.stop_playlist()
            return

        # Safety check for current case limit instead of global
        current_case = getattr(self.macro, 'active_case', 'N')
        limit = self.settings.settings_dict["Special_Cases"][current_case].get("Loop_Limit")
        
        if limit and self.loops_done >= limit:
            print(f"Case {current_case} limit {limit} reached. Stopping.")
            if current_case == "N":
                self.stop_playlist()
            else:
                # If a special case finished, return to N
                self.macro.return_to_normal_loop()
            return
        
        # If a special case finished and no limit or limit not reached, but it's not infinite loop...
        # Wait, user said "if the loops arent checked or no limit is set, they will end and return back to the normal loop."
        # This implies Special Cases (SP1, SP2) only loop if THEIR specific loop limit is set or? 
        # Actually usually macros play once. 
        # "if the loopsarent checked or no limit is set, they will end and return back to the normal loop."
        # I'll check if the main "Loop" checkbox applies to SP1/SP2 too? 
        # Probably easier to stick to: SP1 plays once unless Loop_Limit > 1 or something.
        
        if current_case != "N":
            # Check if SP1/SP2 should loop
            # For now, let's assume they return to N after one completion unless we add a specific "Loop" check for them.
            # The user said "tripled ... all tick boxes". 
            # I didn't add a "Loop" checkbox for SP1/SP2 yet, only "Image Stop" and "Alarm".
            # I'll add a "Loop" checkbox for them if needed, but let's follow the plan.
            self.macro.return_to_normal_loop()
            return

        delay = self.settings.settings_dict["Playback"]["Repeat"].get("Delay", 0)
        
        # Apply global loop interval if set (overrides the default delay)
        if self.global_loop_interval > 0:
            delay_ms = self.global_loop_interval
        else:
            delay_ms = int(delay * 1000) if delay > 0 else 1
        
        if self.is_playlist_playing:
            # Play next script in playlist
            self.current_playlist_index += 1
            # Prevent multiple concurrent triggers if events race
            if hasattr(self, '_playlist_timer'):
                print(f"Logic: Cancelling existing playlist timer {self._playlist_timer}")
                self.after_cancel(self._playlist_timer)
            self._playlist_timer = self.after(delay_ms, self.play_next_script_in_playlist)
            print(f"Logic: Scheduled next playlist script in {delay_ms}ms (ID: {self._playlist_timer})")
        elif self.loop_var.get():
             # Single script loop
             if hasattr(self, '_playlist_timer'):
                 print(f"Logic: Cancelling existing loop timer {self._playlist_timer}")
                 self.after_cancel(self._playlist_timer)
             self._playlist_timer = self.after(delay_ms, self.macro.start_playback)
             print(f"Logic: Scheduled loop restart in {delay_ms}ms (ID: {self._playlist_timer})")
        elif self.loop_scripts_var.get():
             # User enabled random loop but started a single script. 
             # We should probably initialize the playlist now?
             # For now, let's only trigger the full loop playlist if they start it via a specific action or if we modify "Play" to respect this flag.
             # Based on request: "it loop through all the scripts available ... "
             # Code logic: If flag is on, Play button will trigger playlist logic.
             pass

    def start_playlist_playback(self):
        """Start playing all scripts in random order"""
        import random
        # Get all scripts
        scripts_path = self.script_listbox.scripts_path
        script_extensions = ['*.pmr'] # Only supported formats for auto-play
        all_scripts = []
        for extension in script_extensions:
            all_found = glob.glob(os.path.join(scripts_path, extension))
            # Filter out special-case macros from normal playlist
            _SP_MACROS = {"sp1.pmr", "sp2.pmr", "sp3.pmr", "sp4.pmr", "sp5.pmr", "sp6.pmr"}
            all_scripts.extend([f for f in all_found if os.path.basename(f).lower() not in _SP_MACROS])
        
        if not all_scripts:
            messagebox.showinfo("No Scripts", "No playable scripts found.")
            return

        self.playlist = random.sample(all_scripts, len(all_scripts))
        self.current_playlist_index = 0
        self.is_playlist_playing = True
        
        # Show overlay
        if self.overlay:
            self.overlay.destroy()
        
        # Get loop limit for Case N
        limit = self.settings.settings_dict["Special_Cases"]["N"].get("Loop_Limit")
        limit_text = f"\nEnds at: {limit}" if limit and int(limit) > 0 else ""
        
        first_script_name = os.path.basename(self.playlist[0])
        self.overlay = Overlay(f"{first_script_name}\nTotal Run: {self.total_playlist_runs}{limit_text}")
        
        self.play_next_script_in_playlist()

    def play_next_script_in_playlist(self):
        """Play the next script in the queue"""
        if not self.is_playlist_playing:
            return

        if not self.playlist or self.current_playlist_index >= len(self.playlist):
            # Refresh and restart
            scripts_path = self.script_listbox.scripts_path
            script_extensions = ['*.pmr']
            all_scripts = []
            for extension in script_extensions:
                all_found = glob.glob(os.path.join(scripts_path, extension))
                _SP_MACROS = {"sp1.pmr", "sp2.pmr", "sp3.pmr", "sp4.pmr", "sp5.pmr", "sp6.pmr"}
                all_scripts.extend([f for f in all_found if os.path.basename(f).lower() not in _SP_MACROS])
            
            if not all_scripts:
                self.stop_playlist()
                return

            self.playlist = random.sample(all_scripts, len(all_scripts))
            self.current_playlist_index = 0
        
        current_script = self.playlist[self.current_playlist_index]
        script_name = os.path.basename(current_script)
        
        # Check if we are resuming an interrupted script (Case N)
        # If so, the previous count for this script shouldn't stand, so we remove it.
        # This ensures "halted n case shouldn't add count".
        if getattr(self.macro, 'case_n_interrupted', False):
            self.total_playlist_runs -= 1
            self.macro.case_n_interrupted = False # Reset flag
            
        self.total_playlist_runs += 1
        
        # Update overlay
        if self.overlay:
            # Get loop limit for Case N (refresh in case it changed)
            limit = self.settings.settings_dict["Special_Cases"]["N"].get("Loop_Limit")
            limit_text = f"\nEnds at: {limit}" if limit and int(limit) > 0 else ""
            self.overlay.update_text(f"{script_name}\nTotal Run: {self.total_playlist_runs}{limit_text}")
        
        try:
            with open(current_script, 'r') as record:
                loaded_content = load(record)
            
            self.macro.import_record(loaded_content)
            self.macro.start_playback()
            print(f"Playing playlist item: {script_name}")
            
        except Exception as e:
            print(f"Error playing script {script_name}: {e}")
            # Skip to next
            self.on_playback_finished()

    def stop_playlist(self):
        self.is_playlist_playing = False
        
        # End Telegram notification session when playlist stops
        # CRITICAL: Do NOT send "Ended Script" if we are just switching cases (SP1/SP2)
        if hasattr(self.macro, 'session_active') and self.macro.session_active:
            if not getattr(self.macro, 'switching_case', False):
                try:
                    if self.telegram_notifier.is_enabled():
                        self.telegram_notifier.send_message("Ended Script")
                        self.macro.session_active = False
                except Exception as e:
                    print(f"Failed to send TG end notification from stop_playlist: {e}")
            else:
                print("Skipping 'Ended Script' notification in stop_playlist (Switching Case)")
        
        if hasattr(self, '_playlist_timer'):
            print(f"Logic: stop_playlist cancelling timer {self._playlist_timer}")
            self.after_cancel(self._playlist_timer)
            delattr(self, '_playlist_timer')
            
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None

    def load_script_folder(self):
        """Open dialog to select script folder and refresh list"""
        folder_selected = filedialog.askdirectory(title="Select Scripts Folder")
        if folder_selected:
            # Update path in script listbox
            self.script_listbox.scripts_path = folder_selected
            self.script_listbox.refresh_script_list()
            # Update title
            self.script_listbox.title_label.config(text=f"Scripts: {os.path.basename(folder_selected)}")
            messagebox.showinfo("Folder Loaded", f"Loaded scripts from:\n{folder_selected}")
            
            # Save Last Folder
            self.settings.change_settings("Others", "Last_Script_Folder", None, folder_selected)
            
            # Load batch settings (Area, Image)
            self.load_batch_settings(folder_selected)

    def load_batch_settings(self, folder_path):
        """Load specific settings for this folder (batch)"""
        config_path = os.path.join(folder_path, "batch_config.json")
        
        # Image files per case
        img_name_map = {
            "N": "target.png",
            "SP1": "sp1target.png",
            "SP2": "sp2target.png",
            "SP3": "sp3target.png",
            "SP4": "sp4target.png",
            "SP5": "sp5target.png",
            "SP6": "sp6target.png",
        }
        for case, img_name in img_name_map.items():
            target_image = os.path.join(folder_path, img_name)
            if os.path.exists(target_image):
                self.settings.change_settings("Special_Cases", case, "Image_Path", target_image)
                print(f"Loaded batch image for {case}: {target_image}")
        
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    
                    # Areas — stored as list-of-lists; migrate old single-list format
                    areas = config.get("Areas", {})
                    if isinstance(areas, dict):
                        for case, area in areas.items():
                            if case in ["N", "SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
                                # Migrate old [x1,y1,x2,y2] → [[x1,y1,x2,y2]]
                                if area and isinstance(area, list) and len(area) == 4 and not isinstance(area[0], list):
                                    area = [area]
                                self.settings.change_settings("Special_Cases", case, "Area", area)
                    else:
                        area = config.get("Area")
                        if area:
                            if isinstance(area, list) and len(area) == 4 and not isinstance(area[0], list):
                                area = [area]
                            self.settings.change_settings("Special_Cases", "N", "Area", area)
                    
                    # TG Alerts (per case toggles)
                    tg_alerts = config.get("TG_Alerts", {})
                    if isinstance(tg_alerts, dict):
                        for case, enabled in tg_alerts.items():
                             if case in ["N", "SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
                                 self.settings.change_settings("Special_Cases", case, "TG_Alert", enabled)

                    # TG Messages
                    tg_msgs = config.get("TG_Messages", {})
                    if isinstance(tg_msgs, dict):
                        for case, msg in tg_msgs.items():
                             if case in ["N", "SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
                                 self.settings.change_settings("Special_Cases", case, "TG_Message", msg)

                    # Telegram Settings - Only load if GLOBAL is missing (User request: "unchanged across folders")
                    current_token = self.settings.settings_dict.get("Others", {}).get("Telegram_Token", "")
                    saved_token = config.get("Telegram_Token", "")
                    saved_chat_id = config.get("Telegram_Chat_ID", "")
                    
                    if not current_token and saved_token and saved_chat_id:
                         self.settings.change_settings("Others", "Telegram_Token", None, saved_token)
                         self.settings.change_settings("Others", "Telegram_Chat_ID", None, saved_chat_id)
                         if saved_token and saved_chat_id:
                             self.update_tg_status_icon(True)

            except Exception as e:
                print(f"Error loading batch config: {e}")
            
        self.sync_ui_with_settings()

    def save_batch_settings(self, case=None):
        """Save current relevant settings to the batch folder"""
        folder_path = self.script_listbox.scripts_path
        if not folder_path or not os.path.exists(folder_path):
            return

        config_path = os.path.join(folder_path, "batch_config.json")
        
        # Load existing config to not overwrite other cases
        data = {"Areas": {}}
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r') as f:
                    data = json.load(f)
                    if "Areas" not in data: data["Areas"] = {}
            except: pass
            
        for c in ["N", "SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
            areas = self.settings.settings_dict["Special_Cases"][c].get("Area")
            if areas:
                # Ensure stored as list-of-lists
                if isinstance(areas, list) and len(areas) == 4 and not isinstance(areas[0], list):
                    areas = [areas]
                data["Areas"][c] = areas
            
            # Also save TG_Alert preference
            data.setdefault("TG_Alerts", {})[c] = self.settings.settings_dict["Special_Cases"][c].get("TG_Alert", False)
            data.setdefault("TG_Messages", {})[c] = self.settings.settings_dict["Special_Cases"][c].get("TG_Message", "")
        
        # User request: Add Telegram data to folder config
        data["Telegram_Token"] = self.settings.settings_dict.get("Others", {}).get("Telegram_Token", "")
        data["Telegram_Chat_ID"] = self.settings.settings_dict.get("Others", {}).get("Telegram_Chat_ID", "")

        try:
            with open(config_path, 'w') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving batch settings: {e}")

    def adjust_confidence(self, delta, case="N"):
        """Increase or decrease image recognition confidence"""
        current = self.settings.settings_dict["Special_Cases"][case].get("Confidence", 0.75)
        new_val = round(max(0.60, min(1.0, current + delta)), 2)
        self.settings.change_settings("Special_Cases", case, "Confidence", new_val)
        self.case_vars[case]["confidence"] = new_val
        if "conf_label" in self.case_vars[case]:
            self.case_vars[case]["conf_label"].config(text=f"{int(new_val * 100)}%")

    def adjust_loop_interval(self, delta):
        """Increase or decrease loop interval (delay)"""
        current = self.settings.settings_dict["Playback"]["Repeat"].get("Delay", 0)
        new_val = max(0, min(10, current + delta))
        self.settings.change_settings("Playback", "Repeat", "Delay", new_val)
        self.int_val_label.config(text=f"{new_val} seconds")

    def _on_canvas_configure(self, event):
        """Update the width of the canvas window to match the canvas"""
        self.main_canvas.itemconfig(self.canvas_window, width=event.width)

class Overlay(Toplevel):
    def __init__(self, text):
        super().__init__()
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        self.geometry(f"+{self.winfo_screenwidth() - 200}+50") # Top Right
        
        self.label = Label(self, text=text, font=("Arial", 12, "bold"), background="yellow")
        self.label.pack(ipadx=10, ipady=5)
        
        # Transparent background trick for windows if needed, but simple yellow box is visible
        
    def update_text(self, text):
        self.label.config(text=text)