import cv2
import numpy as np
from PIL import ImageGrab
import time
from threading import Event

def monitor_process(stop_event: Event, found_event: Event, area: list, image_path: str, confidence: float = 0.75, log_callback=None):
    """
    Monitor a specific screen area for a target image.
    
    Args:
        stop_event: Event to signal this process to stop.
        found_event: Event to signal that the image was found.
        area: List [x1, y1, x2, y2] defining the screen area to monitor.
        image_path: Path to the target image to find.
        confidence: Confidence threshold for detection (0.0 to 1.0).
    """
    if not image_path or not area:
        return

    try:
        # Load the target image
        template = cv2.imread(image_path, cv2.IMREAD_COLOR)
        if template is None:
            print(f"Error: Could not load image from {image_path}")
            return
            
        # Get template dimensions
        h, w = template.shape[:2]
        
        # Calculate width and height of the area
        x1, y1, x2, y2 = area
        area_width = x2 - x1
        area_height = y2 - y1
        
        # Ensure area is valid
        if area_width <= 0 or area_height <= 0:
            print("Error: Invalid area dimensions")
            return

        case_name = "N"
        if "sp1target" in image_path.lower(): case_name = "SP1"
        elif "sp2target" in image_path.lower(): case_name = "SP2"
        elif "sp3target" in image_path.lower(): case_name = "SP3"
        elif "sp4target" in image_path.lower(): case_name = "SP4"
        elif "sp5target" in image_path.lower(): case_name = "SP5"
        elif "sp6target" in image_path.lower(): case_name = "SP6"

        if log_callback: log_callback(f"Monitoring {case_name} (Conf: {confidence})", "debug")
        last_log_time = 0

        while not stop_event.is_set():
            start_time = time.time()
            
            try:
                # Capture the screen area
                # ImageGrab.grab(bbox=(x1, y1, x2, y2))
                screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                
                # Convert PIL image to OpenCV format (BGR)
                screen_np = np.array(screenshot)
                screen_cv = cv2.cvtColor(screen_np, cv2.COLOR_RGB2BGR)
                
                # Perform template matching
                # We use TM_CCOEFF_NORMED for robustness
                res = cv2.matchTemplate(screen_cv, template, cv2.TM_CCOEFF_NORMED)
                
                # Use the provided confidence threshold
                threshold = confidence
                min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
                
                # Periodically log best match (every ~5 seconds or when close)
                current_time = time.time()
                if log_callback and (current_time - last_log_time > 5.0 or max_val > confidence - 0.1):
                    log_callback(f"Case {case_name} match: {int(max_val*100)}% (Goal: {int(confidence*100)}%)", "debug")
                    last_log_time = current_time

                loc = np.where(res >= threshold)
                
                # If we found at least one match
                if max_val >= confidence:
                    if log_callback: log_callback(f"Case {case_name} FOUND! ({int(max_val*100)}%)", "trigger")
                    found_event.set()
                    break
                    
            except Exception as e:
                print(f"Error in monitoring loop: {e}")
                
            # Wait for 0.5 seconds for much more responsive detection
            time.sleep(0.5)
            
    except Exception as e:
        print(f"Fatal error in monitor process: {e}")
    finally:
        print("Monitoring process stopped.")
