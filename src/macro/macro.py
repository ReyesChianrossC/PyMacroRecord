from tkinter import *
from tkinter import messagebox
from pynput import mouse, keyboard
from pynput.mouse import Button
from utils.get_key_pressed import getKeyPressed
from utils.record_file_management import RecordFileManagement
from utils.warning_pop_up_save import confirm_save
from utils.show_toast import show_notification_minim, show_toast
from utils.keys import vk_nb
from time import time, sleep
from os import getlogin, system, path
from sys import platform
from threading import Thread
from datetime import datetime
from json import load
from threading import Thread, Event
from utils.image_monitor import monitor_process
from utils.sound_generator import play_beep
from utils.debug_logger import debug_logger
import winsound

# Global lookup dictionary for special keys mapping
LOOKUP_SPECIAL_KEY = {}

def setup_special_key_mappings():
    """Set up the lookup dictionary for special keyboard keys."""
    global LOOKUP_SPECIAL_KEY
    
    # Modifier keys
    LOOKUP_SPECIAL_KEY[keyboard.Key.alt] = 'alt'
    LOOKUP_SPECIAL_KEY[keyboard.Key.alt_l] = 'altleft'
    LOOKUP_SPECIAL_KEY[keyboard.Key.alt_r] = 'altright'
    LOOKUP_SPECIAL_KEY[keyboard.Key.alt_gr] = 'altright'
    LOOKUP_SPECIAL_KEY[keyboard.Key.ctrl] = 'ctrlleft'
    LOOKUP_SPECIAL_KEY[keyboard.Key.ctrl_r] = 'ctrlright'
    LOOKUP_SPECIAL_KEY[keyboard.Key.shift] = 'shift_left'
    LOOKUP_SPECIAL_KEY[keyboard.Key.shift_r] = 'shiftright'
    
    # System keys
    LOOKUP_SPECIAL_KEY[keyboard.Key.cmd] = 'winleft'
    LOOKUP_SPECIAL_KEY[keyboard.Key.cmd_r] = 'winright'
    LOOKUP_SPECIAL_KEY[keyboard.Key.caps_lock] = 'capslock'
    LOOKUP_SPECIAL_KEY[keyboard.Key.num_lock] = 'num_lock'
    LOOKUP_SPECIAL_KEY[keyboard.Key.scroll_lock] = 'scroll_lock'
    
    # Navigation keys
    LOOKUP_SPECIAL_KEY[keyboard.Key.up] = 'up'
    LOOKUP_SPECIAL_KEY[keyboard.Key.down] = 'down'
    LOOKUP_SPECIAL_KEY[keyboard.Key.left] = 'left'
    LOOKUP_SPECIAL_KEY[keyboard.Key.right] = 'right'
    LOOKUP_SPECIAL_KEY[keyboard.Key.home] = 'home'
    LOOKUP_SPECIAL_KEY[keyboard.Key.end] = 'end'
    LOOKUP_SPECIAL_KEY[keyboard.Key.page_up] = 'pageup'
    LOOKUP_SPECIAL_KEY[keyboard.Key.page_down] = 'pagedown'
    
    # Function keys
    LOOKUP_SPECIAL_KEY[keyboard.Key.f1] = 'f1'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f2] = 'f2'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f3] = 'f3'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f4] = 'f4'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f5] = 'f5'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f6] = 'f6'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f7] = 'f7'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f8] = 'f8'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f9] = 'f9'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f10] = 'f10'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f11] = 'f11'
    LOOKUP_SPECIAL_KEY[keyboard.Key.f12] = 'f12'
    
    # Action keys
    LOOKUP_SPECIAL_KEY[keyboard.Key.enter] = 'enter'
    LOOKUP_SPECIAL_KEY[keyboard.Key.space] = 'space'
    LOOKUP_SPECIAL_KEY[keyboard.Key.tab] = 'tab'
    LOOKUP_SPECIAL_KEY[keyboard.Key.backspace] = 'backspace'
    LOOKUP_SPECIAL_KEY[keyboard.Key.delete] = 'delete'
    LOOKUP_SPECIAL_KEY[keyboard.Key.esc] = 'esc'
    LOOKUP_SPECIAL_KEY[keyboard.Key.insert] = 'insert'
    LOOKUP_SPECIAL_KEY[keyboard.Key.pause] = 'pause'
    LOOKUP_SPECIAL_KEY[keyboard.Key.print_screen] = 'print_screen'
    
    # Media keys
    LOOKUP_SPECIAL_KEY[keyboard.Key.media_play_pause] = 'playpause'

class Macro:
    """Init a new Macro"""


    def __init__(self, main_app):
        self.showEventsOnStatusBar = None
        self.mouseControl = mouse.Controller()
        self.keyboardControl = keyboard.Controller()
        self.record = False
        self.playback = False
        self.macro_events = {}
        self.main_app = main_app
        self.user_settings = self.main_app.settings
        self.main_menu = self.main_app.menu
        self.macro_file_management = RecordFileManagement(self.main_app, self.main_menu)
        
        # Initialize special key mappings
        setup_special_key_mappings()

        self.mouseBeingListened = None
        self.keyboardBeingListened = None
        self.keyboard_listener = None
        self.mouse_listener = None
        self.time = None
        self.event_delta_time=0

        self.keyboard_listener = keyboard.Listener(
                on_press=self.__on_press, on_release=self.__on_release
            )

        self.keyboard_listener.start()

        self.active_case = "N"
        self.hard_stop_triggered = False
        self.monitor_generation = 0 # Track playback sessions
        self.playback_finished = Event()
        self.playback_finished.set()  # Initially set (no playback running)
        
        # Session tracking for Telegram notifications

        self.session_active = False  # Track if we're in an active playback session
        self.switching_case = False  # Track if we are switching cases (don't end session)
        self.case_n_interrupted = False  # Track if Case N was interrupted mid-execution
        
        # Load persisted state if available
        self.load_playback_state()
        
        # Image Recognition Process (Legacy vars - keep for safety)

    def save_playback_state(self):
        """Save current playback state to settings for persistence"""
        try:
            self.user_settings.change_settings("Playback_State", "loops_done", None, self.main_app.loops_done)
            self.user_settings.change_settings("Playback_State", "active_case", None, getattr(self, "active_case", "N"))
        except Exception as e:
            print(f"Error saving playback state: {e}")

    def load_playback_state(self):
        """Load playback state from settings"""
        try:
            state = self.user_settings.settings_dict.get("Playback_State", {})
            # Only restore if it makes sense (don't restore stale state from days ago)
            # For now, we'll just use it if available
            if state:
                self.main_app.loops_done = state.get("loops_done", 0)
                self.active_case = state.get("active_case", "N")
        except Exception as e:
            print(f"Error loading playback state: {e}")


    def center_cursor(self):
        """Move the cursor to the center of the primary screen."""
        if platform.lower() == "win32":
            import ctypes
            user32 = ctypes.windll.user32
            screen_width = user32.GetSystemMetrics(0)
            screen_height = user32.GetSystemMetrics(1)
            center_x = screen_width // 2
            center_y = screen_height // 2
            self.mouseControl.position = (center_x, center_y)

    def start_record(self, by_hotkey=False):
        if self.main_app.prevent_record or self.record:
            return
        if not by_hotkey:
            if not self.main_app.macro_saved and self.main_app.macro_recorded:
                wantToSave = confirm_save(self.main_app)
                if wantToSave:
                    self.macro_file_management.save_macro()
                elif wantToSave is None:
                    return
        self.macro_events = {"events": []}
        self.record = True
        self.time = time()
        self.center_cursor()
        self.event_delta_time=0
        userSettings = self.user_settings.settings_dict
        self.showEventsOnStatusBar = userSettings["Recordings"]["Show_Events_On_Status_Bar"]
        if (
            userSettings["Recordings"]["Mouse_Move"]
            and userSettings["Recordings"]["Mouse_Click"]
        ):
            self.mouse_listener = mouse.Listener(
                on_move=self.__on_move,
                on_click=self.__on_click,
                on_scroll=self.__on_scroll,
            )
            self.mouse_listener.start()
            self.mouseBeingListened = True
        elif userSettings["Recordings"]["Mouse_Move"]:
            self.mouse_listener = mouse.Listener(
                on_move=self.__on_move, on_scroll=self.__on_scroll
            )
            self.mouse_listener.start()
            self.mouseBeingListened = True
        elif userSettings["Recordings"]["Mouse_Click"]:
            self.mouse_listener = mouse.Listener(
                on_click=self.__on_click, on_scroll=self.__on_scroll
            )
            self.mouse_listener.start()
            self.mouseBeingListened = True
        if userSettings["Recordings"]["Keyboard"]:
            self.keyboardBeingListened = True
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["load_text"], state=DISABLED)
        self.main_app.recordBtn.configure(
            image=self.main_app.stopImg, command=self.stop_record
        )
        self.main_app.playBtn.configure(state=DISABLED)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["save_text"], state=DISABLED)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["save_as_text"], state=DISABLED)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["new_text"], state=DISABLED)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["load_text"], state=DISABLED)
        if userSettings["Minimization"]["When_Recording"]:
            self.main_app.withdraw()
            Thread(target=lambda: show_notification_minim(self.main_app)).start()

    def stop_record(self):
        if not self.record:
            return
        userSettings = self.user_settings.settings_dict
        self.record = False
        if self.mouseBeingListened:
            self.mouse_listener.stop()
            self.mouseBeingListened = False
        if self.keyboardBeingListened:
            self.keyboardBeingListened = False
        self.main_app.recordBtn.configure(
            image=self.main_app.recordImg, command=self.start_record
        )
        self.main_app.playBtn.configure(state=NORMAL, command=self.main_app.on_play_click)
        self.main_menu.file_menu.entryconfig(
            self.main_app.text_content["file_menu"]["save_text"], state=NORMAL, command=self.macro_file_management.save_macro
        )
        self.main_menu.file_menu.entryconfig(
            self.main_app.text_content["file_menu"]["save_as_text"], state=NORMAL, command=self.macro_file_management.save_macro_as
        )
        self.main_menu.file_menu.entryconfig(
            self.main_app.text_content["file_menu"]["new_text"], state=NORMAL, command=self.macro_file_management.new_macro
        )
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["load_text"], state=NORMAL)

        self.main_app.macro_recorded = True
        self.main_app.macro_saved = False

        if userSettings["Minimization"]["When_Recording"]:
            self.main_app.deiconify()

    def start_playback(self):
        """Start macro playback"""
        debug_logger.separator()
        debug_logger.log(f"START_PLAYBACK called: active_case={getattr(self, 'active_case', 'N')}, loops_done={self.main_app.loops_done}")
        if self.playback or self.record or getattr(self, "hard_stop_triggered", False):
            return
        
        # --- HARDCODED TELEGRAM NOTIFICATION: Starting Script ---
        # Only send notification when starting a NEW session, not on every loop
        if not self.session_active:
            try:
                if hasattr(self.main_app, 'telegram_notifier') and self.main_app.telegram_notifier.is_enabled():
                    self.main_app.telegram_notifier.send_message("Starting Script")
                    self.session_active = True  # Mark session as active
            except Exception as e:
                print(f"Failed to send TG start notification: {e}")
        # --------------------------------------------------------
        
        self.main_app.log(f"Starting playback (Case {getattr(self, 'active_case', 'N')})", "trigger")
        # self.hard_stop_triggered = False # Removed from here, moved to manual trigger in main_app
        self.manual_stop = False # Reset manual stop flag
        self.playback_finished.clear()  # Mark that playback is starting
        self.monitor_generation += 1 # New generation for this playback
        current_gen = self.monitor_generation
        
        userSettings = self.user_settings.settings_dict
        self.playback = True
        self.center_cursor()
        self.main_app.playBtn.configure(
            image=self.main_app.stopImg, command=lambda: self.stop_playback(True)
        )
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["save_text"], state=DISABLED)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["save_as_text"], state=DISABLED)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["new_text"], state=DISABLED)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["load_text"], state=DISABLED)
        self.main_app.recordBtn.configure(state=DISABLED)
        if userSettings.get("Minimization", {}).get("When_Playing", False):
            self.main_app.withdraw()
            Thread(target=lambda: show_notification_minim(self.main_app)).start()
        
        # Start Image Monitoring for ALL enabled cases
        # Case N has PRIORITY: if its stop fires, hard_stop_triggered is set immediately
        # which causes any racing SP watch threads to abort their interrupt attempt.
        self.active_monitors = {}
        all_cases = ["N", "SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]
        active_case = getattr(self, "active_case", "N")
        for case in all_cases:
            # CRITICAL FIX: When ANY SP case is active, only monitor Case N.
            # This prevents SP cases from cross-interrupting each other
            # (the root cause of the SP2+ infinite loop bug).
            if active_case != "N" and case != "N":
                self.main_app.log(f"Skipping monitor for {case} (SP case {active_case} is active)", "debug")
                continue

            case_settings = userSettings.get("Special_Cases", {}).get(case, {})
            if case_settings.get("Enabled", False):
                image_path = case_settings.get("Image_Path")
                areas = case_settings.get("Area")  # now a list of [x1,y1,x2,y2]
                confidence = case_settings.get("Confidence", 0.75)
                
                # Validate image path
                if not image_path or not path.exists(image_path):
                    self.main_app.log(f"Skipping monitor for Case {case} - image not found: {image_path}", "debug")
                    continue

                # Normalise area to list-of-lists
                if not areas:
                    self.main_app.log(f"Skipping monitor for Case {case} - no area defined", "debug")
                    continue
                if isinstance(areas, list) and len(areas) == 4 and not isinstance(areas[0], list):
                    areas = [areas]  # migrate old single-area format

                # Spawn one monitor thread per area
                for idx, area in enumerate(areas):
                    if not area or len(area) != 4:
                        continue
                    monitor_key = f"{case}_{idx}"
                    self.main_app.log(f"Starting monitor for Case {case} area #{idx}...", "debug")
                    found_evt = Event()
                    stop_evt  = Event()

                    mon_thread = Thread(
                        target=monitor_process,
                        args=(stop_evt, found_evt, area, image_path, confidence),
                        kwargs={"log_callback": self.main_app.log},
                        daemon=True
                    )
                    mon_thread.start()

                    self.active_monitors[monitor_key] = {
                        "case": case,
                        "proc": mon_thread,
                        "found_evt": found_evt,
                        "stop_evt": stop_evt
                    }
                    # Watch thread passes the CASE name so behaviour is unchanged
                    Thread(target=self.__watch_monitor, args=(monitor_key, case, current_gen), daemon=True).start()
                self.main_app.log(f"Monitoring Case {case} ({len(areas)} area(s))...", "debug")

        if userSettings["Playback"].get("Loop_Scripts", False):
            Thread(target=self.__play_events).start()
        elif userSettings["Playback"]["Repeat"]["Interval"] > 0:
            Thread(target=self.__play_interval).start()
        elif userSettings["Playback"]["Repeat"]["For"] > 0:
            Thread(target=self.__play_for).start()
        else:
            Thread(target=self.__play_events).start()

        # if getattr(self, "active_case", "N") == "N":
        #     show_toast("Macro playback started.", title="Case N Starting")
        
        print(f"playback started for Case {getattr(self, 'active_case', 'N')}")

    def __watch_monitor(self, monitor_key, case, generation):
        """Watch for the found event from a specific monitor thread.
        monitor_key  : key into self.active_monitors (e.g. 'SP1_0')
        case         : the logical case name ('N', 'SP1', 'SP2', 'SP3', 'SP4', 'SP5', 'SP6')
        generation   : playback-session ID to reject zombie notifications
        """
        mon = self.active_monitors.get(monitor_key)
        if mon and mon["found_evt"]:
            mon["found_evt"].wait()
            
            # CRITICAL: Check if this is a "zombie" monitor from a previous session
            if generation != self.monitor_generation:
                return

            # CRITICAL: Check if a stop was already triggered by another thread while we were waiting
            if getattr(self, "hard_stop_triggered", False) or getattr(self, "manual_stop", False):
                return

            if self.playback:
                print(f"Target image for Case {case} found! (area key: {monitor_key})")
                
                # Get case settings
                userSettings = self.user_settings.settings_dict
                case_settings = userSettings.get("Special_Cases", {}).get(case, {})
                
                # Play alarm FIRST if enabled (for ANY case)
                if case_settings.get("Alarm", False):
                    Thread(target=self.__play_alarm).start()
                
                # --- TELEGRAM NOTIFICATION HOOK ---
                try:
                    self.main_app.send_telegram_alert(case)
                except Exception as e:
                    print(f"Failed to send TG alert: {e}")
                # ----------------------------------

                stop_program = case_settings.get("Stop_Program", False)
                debug_logger.log(f"MONITOR HIT: Case={case}, key={monitor_key}, Stop_Program={stop_program}, Active={self.active_case}")

                if case == "N" or stop_program:
                    # ── STOP BEHAVIOUR (Case N is absolute priority) ──────────
                    # SET FLAGS IMMEDIATELY to block ALL other threads (including
                    # any SP watch threads that might race to interrupt).
                    self.hard_stop_triggered = True
                    self.manual_stop = True

                    log_msg = f"CASE {case} IMAGE DETECTED - STOPPING ALL (Stop_Prog={stop_program})"
                    self.main_app.log(log_msg, "stop")
                    self.main_app.after(0, self.stop_playback, True)
                else:
                    # ── INTERRUPT BEHAVIOUR (SP1-SP4 without Stop Program) ────
                    # If Case N already fired, abort — N has absolute priority.
                    if getattr(self, "hard_stop_triggered", False):
                        debug_logger.log(f"Interrupt for {case} skipped: Hard Stop already active")
                        return

                    self.main_app.log(f"CASE {case} IMAGE DETECTED - INTERRUPTING", "trigger")
                    self.main_app.after(0, lambda: self.interrupt_for_special_case(case))

    def interrupt_for_special_case(self, case):
        """Stop current playback and start special case macro (Threaded)"""
        # Run the heavy switching logic in a separate thread to avoid freezing UI
        Thread(target=self._switch_case_thread, args=(case,)).start()

    def _switch_case_thread(self, case):
        """Background thread to handle switching cases"""
        if getattr(self, "hard_stop_triggered", False):
            self.main_app.log(f"Interrupt for {case} cancelled (Stop Program priority)", "debug")
            return

        self.main_app.log(f"Switching to Case {case}...", "info")
        
        # Mark that Case N was interrupted (if currently playing Case N)
        if getattr(self, 'active_case', 'N') == 'N':
            self.case_n_interrupted = True
            debug_logger.log("Case N interrupted - loop count will not increment")
        
        # CRITICAL: Stop ALL area monitor threads for this special case to prevent re-triggering
        for key, mon in list(self.active_monitors.items()):
            if mon.get("case") == case:
                if mon["stop_evt"]:
                    mon["stop_evt"].set()
                self.main_app.log(f"Stopped monitor {key} ({case} triggered)", "debug")
        
        # Stop playback (this is fast, sets flags)
        # Use switching_case flag to prevent "Ended Script" notification
        self.switching_case = True
        self.stop_playback(True) 
        self.switching_case = False
        
        # CRITICAL: Wait for the old playback thread to fully terminate
        # unique_id = time()
        self.main_app.log(f"Waiting for old playback to finish...", "debug")
        if not self.playback_finished.wait(timeout=3.0):
             self.main_app.log("Warning: Old playback timed out. Forcing switch.", "info")
        
        # CRITICAL: Check for hard stop again after waiting for playback to finish
        if getattr(self, "hard_stop_triggered", False):
            self.main_app.log(f"Case switch to {case} aborted (Hard Stop active)", "stop")
            return

        self.main_app.log(f"Old playback finished, checking {case} macro", "debug")
        
        # Try several common locations for the macro
        current_folder = self.main_app.script_listbox.scripts_path
        macro_name = case.lower() + ".pmr"
        
        # Priority 1: Current scripts folder
        macro_path = path.join(current_folder, macro_name)
        
        # Priority 2: If not found, check the base scripts folder
        base_scripts = path.abspath(path.join(path.dirname(__file__), "..", "scripts"))
        if not path.exists(macro_path):
             macro_path = path.join(base_scripts, macro_name)
        
        # NEW: Gracefully handle missing macro files for SP1/SP2
        if not path.exists(macro_path):
            self.main_app.log(f"No macro file found for {case}, detection-only mode", "info")
            
            # Detection worked but no macro to play - just resume normal flow
            # The alarm/stop already happened in __watch_monitor before we got here
            debug_logger.log(f"No macro for {case}, switching to N (loops_done={self.main_app.loops_done})")
            self.active_case = "N"
            # Keep loops_done to continue counter across interruptions
            self.main_app.log(f"{case} detection completed (no macro), resuming normal loop", "info")
            
            # Resume normal playback if not hard stopped
            if not getattr(self, "hard_stop_triggered", False):
                self.main_app.after(100, lambda: self.main_app.on_play_click(manual_start=False))
            return

        try:
            with open(macro_path, 'r') as f:
                loaded_content = load(f)
            
            # show_toast(f"Detecting image... Triggering Case {case}!", title="Special Case Triggered")
            # active_case is a string, safe.
            debug_logger.log(f"Switching to Case {case} (loops_done={self.main_app.loops_done} PRESERVED)")
            self.active_case = case
            # Keep loops_done to continue counter across interruptions
            self.import_record(loaded_content)
            
            # Now trigger playback on main thread
            self.main_app.after(50, self.start_playback)
            
        except Exception as e:
            print(f"Error loading {case} macro from {macro_path}: {e}")
            self.main_app.log(f"Error loading {case}: {e}", "stop")
            self.active_case = "N"
            self.main_app.after(50, self.start_playback)
            
        except Exception as e:
            print(f"Error loading {case} macro from {macro_path}: {e}")
            self.active_case = "N"
            self.main_app.after(50, self.start_playback)

    def return_to_normal_loop(self):
        """Called when SP1/SP2 finishes to return to Case N"""
        finished_case = getattr(self, 'active_case', 'UNKNOWN')
        debug_logger.log(f"return_to_normal_loop called. Finished: {finished_case}")
        # show_toast(f"Special Case {finished_case} finished. Resuming Case N.", title="Resuming Normal Loop")
        
        # We don't need to force stop_playback(True) here if it was already stopped naturally.
        # Calling it might interfere with playlist state or set manual_stop unnecessarily.
        # But we ensures clean state.
        if self.playback:
            self.switching_case = True
            self.stop_playback(True)
            self.switching_case = False
        
        debug_logger.log(f"Return to normal loop from {finished_case} (loops_done={self.main_app.loops_done} PRESERVED)")
        self.active_case = "N"
        # Keep loops_done to continue counter across interruptions
        
        # Restart the normal flow (playlist or single macro)
        # CRITICAL: Do NOT resume if a hard stop was triggered
        if getattr(self, "hard_stop_triggered", False):
            self.main_app.log("Resumption cancelled (Hard Stop active)", "stop")
            return

        self.main_app.log("Resuming normal loop...", "info")
        
        # Apply global loop interval if set
        delay_ms = 200  # Default delay
        if hasattr(self.main_app, 'global_loop_interval') and self.main_app.global_loop_interval > 0:
            delay_ms = self.main_app.global_loop_interval
        
        # Increase delay to ensure threads settle
        self.main_app.after(delay_ms, lambda: self.main_app.on_play_click(manual_start=False))

    def __play_alarm(self):
        """Play high beep 3 times"""
        # Get volume from settings
        userSettings = self.user_settings.settings_dict
        volume = userSettings.get("Image_Recognition", {}).get("Volume", 50)
        
        for _ in range(3):
            try:
                play_beep(2000, 300, volume) # 2000Hz, 300ms, volume
                sleep(0.1)
            except:
                pass


    def __play_interval(self):
        userSettings = self.user_settings.settings_dict
        if userSettings["Playback"]["Repeat"]["For"] > 0:
            self.__play_for()
        else:
            self.__play_events()
        timer = time()
        while self.playback:
            sleep(1)
            if time() - timer >= userSettings["Playback"]["Repeat"]["Interval"]:
                if userSettings["Playback"]["Repeat"]["For"] > 0:
                    self.__play_for()
                else:
                    self.__play_events()
                timer = time()

    def __play_for(self):
        userSettings = self.user_settings.settings_dict
        debut = time()
        while self.playback and (time() - debut) < userSettings["Playback"]["Repeat"]["For"]:
            self.__play_events()
        if userSettings["Playback"]["Repeat"]["Interval"] == 0:
            self.stop_playback()

    def __play_events(self):
        global keyToPress
        userSettings = self.user_settings.settings_dict
        click_func = {
            "leftClickEvent": Button.left,
            "rightClickEvent": Button.right,
            "middleClickEvent": Button.middle,
        }
        keyToUnpress = []

        is_infinite = userSettings["Playback"]["Repeat"].get("Infinite", False)
        # Override infinite loop if Random Loop (Loop_Scripts) is enabled
        if userSettings["Playback"].get("Loop_Scripts", False):
            is_infinite = False

        # CRITICAL: Special Cases (SP1-SP4) should NOT be treated as infinite loops.
        # They play ONCE per trigger.
        active_case = getattr(self, "active_case", "N")
        if active_case in ["SP1", "SP2", "SP3", "SP4", "SP5", "SP6"]:
            is_infinite = False
            userSettings = userSettings.copy() # Shallow copy to not affect global dict
            if "Playback" not in userSettings: userSettings["Playback"] = {}
            if "Repeat" not in userSettings["Playback"]: userSettings["Playback"]["Repeat"] = {}
            # Force 1 time execution for SP cases
            repeat_times = 1
            # Ensure we don't look at For/Scheduled for SP cases
            userSettings["Playback"]["Repeat"]["For"] = 0 
            userSettings["Playback"]["Repeat"]["Scheduled"] = 0
        else:
            if userSettings["Playback"]["Repeat"]["For"] > 0:
                repeat_times = 1
            elif is_infinite:
                repeat_times = float('inf')
            else:
                repeat_times = userSettings["Playback"]["Repeat"]["Times"]

        if userSettings["Playback"]["Repeat"]["Scheduled"] > 0:
            now = datetime.now()
            seconds_since_midnight = (now - now.replace(hour=0, minute=0, second=0, microsecond=0)).total_seconds()
            secondsToWait = userSettings["Playback"]["Repeat"]["Scheduled"] - seconds_since_midnight
            if secondsToWait < 0:
                secondsToWait = 86400 + secondsToWait  # 86400 + -secondsToWait. Meaning it will happen tomorrow
            sleep(secondsToWait)

        repeat_count = 0
        now = time()

        while self.playback and (is_infinite or repeat_count < repeat_times):
            for events in range(len(self.macro_events["events"])):
                elapsed_time = int(time() - now)
                if events % 50 == 0:
                    repeat_display = "∞" if is_infinite else str(repeat_times)
                    self.main_app.status_text.configure(
                        text=f"Repeat: {repeat_count + 1}/{repeat_display}, Time elapsed: {elapsed_time}s")

                if not self.playback:
                    self.unPressEverything(keyToUnpress)
                    return

                if userSettings["Others"]["Fixed_timestamp"] > 0:
                    timeSleep = userSettings["Others"]["Fixed_timestamp"]
                else:
                    timeSleep = (
                            self.macro_events["events"][events]["timestamp"]
                            * (1 / userSettings["Playback"]["Speed"])
                    )
                if timeSleep < 0:
                    timeSleep = abs(timeSleep)
                
                # High Precision Sleep for smooth playback
                # Windows sleep() has 15ms resolution. For small delays, busy-wait.
                if timeSleep > 0.015:
                    sleep(timeSleep)
                elif timeSleep > 0:
                    target_time = time() + timeSleep
                    while time() < target_time:
                        pass # Busy wait for precision
                
                # CRITICAL FIX: Check if playback was stopped during sleep
                if not self.playback:
                    self.unPressEverything(keyToUnpress)
                    return

                event_type = self.macro_events["events"][events]["type"]

                if event_type == "cursorMove":  # Cursor Move
                    x = self.macro_events["events"][events]["x"]
                    y = self.macro_events["events"][events]["y"]
                    # Log removed for smoothness
                    self.mouseControl.position = (x, y)

                elif event_type in click_func:  # Mouse Click
                    x = self.macro_events["events"][events]["x"]
                    y = self.macro_events["events"][events]["y"]
                    pressed = self.macro_events["events"][events]["pressed"]
                    action = "Press" if pressed else "Release"
                    self.main_app.log(f"{action} {event_type} at ({x}, {y})", "debug")
                    self.mouseControl.position = (x, y)
                    if pressed:
                        self.mouseControl.press(click_func[event_type])
                    else:
                        self.mouseControl.release(click_func[event_type])

                elif event_type == "scrollEvent":
                    self.mouseControl.scroll(
                        self.macro_events["events"][events]["dx"],
                        self.macro_events["events"][events]["dy"],
                    )

                elif event_type == "keyboardEvent":  # Keyboard Press,Release
                    if self.macro_events["events"][events]["key"] is not None:
                        try:
                            keyToPress = (
                                self.macro_events["events"][events]["key"]
                                if "Key." not in self.macro_events["events"][events]["key"]
                                else eval(self.macro_events["events"][events]["key"])
                            )
                            if isinstance(keyToPress, str):
                                if ">" in keyToPress:
                                    try:
                                        keyToPress = vk_nb[keyToPress]
                                    except:
                                        keyToPress = None
                            if self.playback:
                                if keyToPress is not None:
                                    if (
                                            self.macro_events["events"][events]["pressed"]
                                            == True
                                    ):
                                        self.keyboardControl.press(keyToPress)
                                        if keyToPress not in keyToUnpress:
                                            keyToUnpress.append(keyToPress)
                                    else:
                                        self.keyboardControl.release(keyToPress)
                        except ValueError as e:
                            if keyToPress is None:
                                pass
                            else:
                                messagebox.showerror("Error",
                                                     f"Error during playback \"{e}\". Please open an issue on Github.")
                            self.stop_playback()
                            return
        
            # Signal that playback has finished
            self.playback_finished.set()
            self.main_app.log("Playback thread terminated", "debug")
            repeat_count += 1
            
            # Only count loops for Case N, not for SP cases
            active_case = getattr(self, 'active_case', 'N')
            if active_case == "N":
                # Only increment if Case N completed fully (not interrupted)
                if not self.case_n_interrupted:
                    old_count = self.main_app.loops_done
                    self.main_app.loops_done += 1
                    debug_logger.log(f"Case N finished: loops_done {old_count} -> {self.main_app.loops_done}")
                    # Save state to persist across interruptions
                    self.save_playback_state()
                else:
                    # Case N was interrupted, reset the flag but don't increment count
                    debug_logger.log(f"Case N was interrupted - loops_done remains {self.main_app.loops_done}")
                    self.case_n_interrupted = False  # Reset for next iteration
            else:
                debug_logger.log(f"Case {active_case} finished: loops_done={self.main_app.loops_done} (NOT incremented)")

            # Hard stop if limit reached
            case_settings = userSettings.get("Special_Cases", {}).get(active_case, {})
            loop_limit = case_settings.get("Loop_Limit")
            
            # CRITICAL FIX: Only enforce loop limit for Case N. 
            # SP1/SP2 don't have their own loop limits in this context, and shouldn't stop based on N's count.
            if active_case == "N" and loop_limit and self.main_app.loops_done >= loop_limit:
                print(f"Logic: Case {active_case} loop limit {loop_limit} reached. Stopping playback.")
                debug_logger.log(f"Loop limit reached for {active_case}: limit={loop_limit}, done={self.main_app.loops_done}")
                self.stop_playback(True)
                return
            elif active_case != "N":
                debug_logger.log(f"Skipping loop limit check for {active_case} (limit={loop_limit}, done={self.main_app.loops_done})")

            if userSettings["Playback"]["Repeat"]["Delay"] > 0:
                if is_infinite or repeat_count < repeat_times:
                    sleep(userSettings["Playback"]["Repeat"]["Delay"])

        self.unPressEverything(keyToUnpress)
        is_loop_scripts = userSettings["Playback"].get("Loop_Scripts", False)
        if (userSettings["Playback"]["Repeat"]["Interval"] == 0 and userSettings["Playback"]["Repeat"]["For"] == 0) or is_loop_scripts:
            # Only stop if we actually completed (not infinite loop)
            if not is_infinite and repeat_count:
                self.stop_playback()

    def unPressEverything(self, keyToUnpress):
        for key in keyToUnpress:
            self.keyboardControl.release(key)
        self.mouseControl.release(Button.left)
        self.mouseControl.release(Button.middle)

    def stop_playback(self, playback_stopped_manually=False):
        debug_logger.log(f"stop_playback called: manually={playback_stopped_manually}")
        self.playback = False
        self.main_app.log("Playback stopped", "stop")
        
        # --- HARDCODED TELEGRAM NOTIFICATION: Ended Script ---
        # Only send notification when session truly ends (not between loops or case switches)
        
        # Check if we're in playlist mode (random loop)
        is_playlist_active = getattr(self.main_app, 'is_playlist_playing', False)
        
        # We are STOPPING the session if:
        # 1. User stopped it manually AND we are NOT just switching cases internally
        # 2. Hard Stop (Image Stop / Stop Program) was triggered
        # 3. Playback finished naturally (no more loops) AND playlist is not continuing
        
        should_end_session = (
            (playback_stopped_manually and not getattr(self, "switching_case", False)) or
            getattr(self, "hard_stop_triggered", False) or
            (not self.playback and not is_playlist_active and not getattr(self, "switching_case", False))
        )
        
        # Don't end session if we're just switching between cases (SP1/SP2 interruption) or between playlist items
        if should_end_session and self.session_active:
            try:
                if hasattr(self.main_app, 'telegram_notifier') and self.main_app.telegram_notifier.is_enabled():
                    self.main_app.telegram_notifier.send_message("Ended Script")
                    self.session_active = False  # Mark session as ended
            except Exception as e:
                print(f"Failed to send TG end notification: {e}")
        # ------------------------------------------------------
        
        # Save state before stopping
        self.save_playback_state()
        
        # Cancel any pending playback timers in the main app
        if hasattr(self.main_app, '_playlist_timer'):
            try:
                self.main_app.after_cancel(self.main_app._playlist_timer)
                delattr(self.main_app, '_playlist_timer')
            except Exception as e:
                pass

        if not playback_stopped_manually:
            pass
        else:
            self.manual_stop = True # Prevent monitors from triggering
            # Potential cause of double app launch on hotkey stop if win10toast is buggy
            # show_toast("Macro playback stopped.", title="Playback Stopped")
            # DISABLED TO PREVENT DUPLICATE LAUNCH ISSUES
            # show_toast("Macro playback stopped.", title="Playback Stopped") 
            
            if hasattr(self.main_app, 'stop_playlist'):
                # Force stop playlist even if active_case is being switched
                # CRITICAL: But only if we are NOT switching cases internally
                if not getattr(self, "switching_case", False):
                    self.main_app.stop_playlist()
                else:
                    debug_logger.log("stop_playlist call skipped (Switching Case)")

        userSettings = self.user_settings.settings_dict
        self.main_app.recordBtn.configure(state=NORMAL)
        self.main_app.playBtn.configure(
            image=self.main_app.playImg, command=self.main_app.on_play_click
        )
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["save_text"], state=NORMAL)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["save_as_text"], state=NORMAL)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["new_text"], state=NORMAL)
        self.main_menu.file_menu.entryconfig(self.main_app.text_content["file_menu"]["load_text"], state=NORMAL)
        if userSettings["Minimization"]["When_Playing"] and playback_stopped_manually:
            self.main_app.deiconify()
        
        # Shutdown logic omitted for brevity in replace, keeping the same logic as before...
        # (Assuming the original lines 561-592 are handled or I should include them)
        # I'll include them to be safe.
        if userSettings["After_Playback"]["Mode"] != "Idle" and not playback_stopped_manually:
            if userSettings["After_Playback"]["Mode"].lower() == "standby":
                if platform == "win32":
                    system("rundll32.exe powrprof.dll, SetSuspendState 0,1,0")
                elif "linux" in platform.lower():
                    system("subprocess.callctl suspend")
                elif "darwin" in platform.lower():
                    system("pmset sleepnow")
            elif userSettings["After_Playback"]["Mode"].lower() == "log off computer":
                if platform == "win32":
                    system("shutdown /l")
                else:
                    system(f"pkill -KILL -u {getlogin()}")
            elif userSettings["After_Playback"]["Mode"].lower() == "turn off computer":
                if platform == "win32":
                    system("shutdown /s /t 0")
                else:
                    system("shutdown -h now")
            elif userSettings["After_Playback"]["Mode"].lower() == "restart computer":
                if platform == "win32":
                    system("shutdown /r /t 0")
                else:
                    system("shutdown -r now")
            elif userSettings["After_Playback"]["Mode"].lower() == "hibernate (if enabled)":
                if platform == "win32":
                    system("shutdown -h")
                elif "linux" in platform.lower():
                    system("systemctl hibernate")
                elif "darwin" in platform.lower():
                    system("pmset sleepnow")
            force_close = True
            self.main_app.quit_software(force_close)

        # Stop ALL Image Monitoring in a separate thread to avoid blocking main thread
        if hasattr(self, 'active_monitors') and self.active_monitors:
            monitors_to_stop = self.active_monitors
            self.active_monitors = {}
            Thread(target=self.__stop_monitors_async, args=(monitors_to_stop,)).start()

        if not playback_stopped_manually and hasattr(self.main_app, 'on_playback_finished'):
            self.main_app.after(0, self.main_app.on_playback_finished)

    def __stop_monitors_async(self, monitors):
        """Clean up monitor threads without blocking"""
        for key, mon in monitors.items():
            try:
                if mon["stop_evt"]:
                    mon["stop_evt"].set()
                if mon["proc"] and mon["proc"].is_alive():
                    mon["proc"].join(timeout=0.2)
            except Exception as e:
                print(f"Error stopping monitor {key}: {e}")


    def import_record(self, record):
        self.main_app.log(f"Importing macro events: {len(record.get('events', []))} events", "debug")
        self.macro_events = record

    def __record_event(self,e):
        e['timestamp'] = self.event_delta_time
        self.macro_events["events"].append(e)

    def __get_event_delta_time(self):
        timenow=time()
        self.event_delta_time = timenow - self.time
        self.time=timenow

    def __on_move(self, x, y):
        self.__get_event_delta_time()
        self.__record_event(
            {"type": "cursorMove", "x": x, "y": y}
        )
        if self.showEventsOnStatusBar:
            self.main_app.status_text.configure(text=f"cursorMove {x} {y}")

    def __on_click(self, x, y, button, pressed):
        self.__get_event_delta_time()
        button_event = "unknownButtonClickEvent"
        if button == Button.left:
            button_event = "leftClickEvent"
        elif button == Button.right:
            button_event = "rightClickEvent"
        elif button == Button.middle:
            button_event = "middleClickEvent"
        self.__record_event(
            {
                "type": button_event,
                "x": x,
                "y": y,
                "pressed": pressed
            }
        )
        if self.showEventsOnStatusBar:
            self.main_app.status_text.configure(text=f"{button_event} {x} {y} {pressed}")

    def __on_scroll(self, x, y, dx, dy):
        self.__get_event_delta_time()
        self.__record_event(
            {"type": "scrollEvent", "dx": dx, "dy": dy}
        )
        if self.showEventsOnStatusBar:
            self.main_app.status_text.configure(text=f"scrollEvent {dx} {dy}")

    def __on_press(self, key):
        self.__get_event_delta_time()
        if self.keyboardBeingListened:
            try:
                # Handle regular character keys
                if hasattr(key, 'char') and key.char:
                    keyPressed = key.char
                else:
                    # Handle special keys using the lookup dictionary
                    keyPressed = LOOKUP_SPECIAL_KEY.get(key, None)
                if keyPressed:
                    self.__record_event(
                        {
                            "type": "keyboardEvent",
                            "key": keyPressed,
                            "pressed": True,
                        }
                    )
                    if self.showEventsOnStatusBar:
                        self.main_app.status_text.configure(text=f"keyboardEvent {keyPressed} pressed")
            except AttributeError:
                pass

    def __on_release(self, key):
        self.__get_event_delta_time()
        if self.keyboardBeingListened:
            try:
                # Handle regular character keys
                if hasattr(key, 'char') and key.char:
                    keyPressed = key.char
                else:
                    # Handle special keys using the lookup dictionary
                    keyPressed = LOOKUP_SPECIAL_KEY.get(key, None)
                if keyPressed:
                    self.__record_event(
                        {
                            "type": "keyboardEvent",
                            "key": keyPressed,
                            "pressed": False,
                        }
                    )
                    if self.showEventsOnStatusBar:
                        self.main_app.status_text.configure(text=f"keyboardEvent {keyPressed} released")
            except AttributeError:
                pass