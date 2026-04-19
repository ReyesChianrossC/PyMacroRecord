
import sys
import os
from unittest.mock import MagicMock, patch

# Correctly mock pynput as a package with submodules
pynput_mock = MagicMock()
sys.modules['pynput'] = pynput_mock
sys.modules['pynput.mouse'] = pynput_mock.mouse
sys.modules['pynput.keyboard'] = pynput_mock.keyboard

sys.modules['win10toast'] = MagicMock()
sys.modules['cv2'] = MagicMock()
sys.modules['PIL'] = MagicMock()
sys.modules['pystray'] = MagicMock()

# Add src to path
sys.path.append(os.path.abspath("."))

# Mock internal modules causing circularity
sys.modules['utils.show_toast'] = MagicMock()
sys.modules['utils.get_key_pressed'] = MagicMock()
sys.modules['utils.record_file_management'] = MagicMock()
sys.modules['utils.warning_pop_up_save'] = MagicMock()
sys.modules['utils.get_file'] = MagicMock()
sys.modules['utils.version'] = MagicMock()
sys.modules['utils.not_windows'] = MagicMock()
sys.modules['windows.popup'] = MagicMock()
sys.modules['windows.main.overlay'] = MagicMock()

from macro.macro import Macro

class Simulation:
    def __init__(self):
        self.scheduled_timers = {}
        self.timer_id_counter = 0

    def after(self, ms, func, *args):
        self.timer_id_counter += 1
        timer_id = f"timer_{self.timer_id_counter}"
        self.scheduled_timers[timer_id] = (func, args)
        print(f"[SIM] Scheduled {func.__name__ if hasattr(func, '__name__') else 'lambda'} in {ms}ms (ID: {timer_id})")
        return timer_id

    def after_cancel(self, timer_id):
        if timer_id in self.scheduled_timers:
            print(f"[SIM] Cancelled timer {timer_id}")
            del self.scheduled_timers[timer_id]
        else:
            print(f"[SIM] Attempted to cancel non-existent timer {timer_id}")

def test_halt_logic():
    print("--- STARTING FINAL HALT LOGIC SIMULATION ---")
    sim = Simulation()
    mock_app = MagicMock()
    mock_app.after = sim.after
    mock_app.after_cancel = sim.after_cancel
    mock_app.script_listbox.scripts_path = "./scripts/Test"
    mock_app.text_content = {"file_menu": {"save_text": "Save", "save_as_text": "Save As", "new_text": "New", "load_text": "Load"}}
    
    mock_settings = MagicMock()
    mock_settings.settings_dict = {
        "Playback": {"Speed": 1, "Repeat": {"Times": 1, "Interval": 0, "For": 0, "Delay": 0.5, "Infinite": True}},
        "Minimization": {"When_Playing": False},
        "After_Playback": {"Mode": "Idle"},
        "Others": {"Fixed_timestamp": 0},
        "Special_Cases": {
            "SP1": {"Enabled": True, "Image_Path": "sp1.png", "Area": [0,0,10,10], "Confidence": 0.8, "Loop_Limit": 1},
            "N": {"Enabled": True, "Image_Path": "n.png", "Area": [0,0,10,10], "Confidence": 0.8}
        }
    }
    
    with patch('macro.macro.keyboard'), patch('macro.macro.mouse'), patch('macro.macro.load') as mock_load:
        mock_load.return_value = {"events": []}
        macro = Macro(mock_app)
        macro.user_settings = mock_settings
        macro.playback = True
        macro.active_case = "N"
        
        mock_app._playlist_timer = sim.after(500, macro.start_playback)
        
        with patch("builtins.open", MagicMock()):
            macro.interrupt_for_special_case("SP1")
                
        print(f"\n[VERIFY] Is timer_1 cancelled? {'timer_1' not in sim.scheduled_timers}")
        assert 'timer_1' not in sim.scheduled_timers
        
        print(f"[VERIFY] Is _playlist_timer attribute gone? {not hasattr(mock_app, '_playlist_timer')}")
        assert not hasattr(mock_app, '_playlist_timer')
        
        print(f"[VERIFY] Is Active Case SP1? {macro.active_case == 'SP1'}")
        assert macro.active_case == "SP1"

    print("\n--- SIMULATION SUCCESSFUL ---")

if __name__ == "__main__":
    test_halt_logic()
