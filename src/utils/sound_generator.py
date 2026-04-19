import numpy as np
import winsound
import io
import wave
import struct
import time
import tempfile
import os

def play_beep(frequency, duration_ms, volume_percent):
    """
    Play a beep sound with a specific frequency, duration, and volume.
    
    :param frequency: Frequency in Hz (e.g., 4400)
    :param duration_ms: Duration in milliseconds (e.g., 500)
    :param volume_percent: Volume from 0 to 100
    """
    if volume_percent <= 0:
        return

    sample_rate = 44100
    duration_s = duration_ms / 1000.0
    volume = volume_percent / 100.0
    
    # Generate sine wave
    t = np.linspace(0, duration_s, int(sample_rate * duration_s), False)
    # Generate samples: amplitude * sin(2 * pi * frequency * t)
    # Amplitude for 16-bit audio is roughly 32767. Max volume uses full range.
    amplitude = 32767 * volume
    audio_data = (amplitude * np.sin(2 * np.pi * frequency * t)).astype(np.int16)
    
    # Convert numpy array to raw bytes (little-endian)
    raw_audio = audio_data.tobytes()
    
    # Save to temporary file for robust async playback
    temp_wav = os.path.join(tempfile.gettempdir(), "pymacro_alarm.wav")
    
    try:
        with wave.open(temp_wav, 'wb') as wave_file:
            wave_file.setnchannels(1)  # Mono
            wave_file.setsampwidth(2)  # 16-bit
            wave_file.setframerate(sample_rate)
            wave_file.writeframes(raw_audio)
        
        # Play sound from FILE ASYNCHRONOUSLY
        # SND_FILENAME = 0x00020000
        # SND_ASYNC = 0x0001
        # SND_NODEFAULT = 0x0002
        flags = winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_NODEFAULT
        winsound.PlaySound(temp_wav, flags)
        
        # Sleep manually to allow sound to play, but release GIL for other threads
        time.sleep(duration_s)
        
    except Exception as e:
        print(f"Error playing sound: {e}")
