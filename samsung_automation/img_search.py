import pyautogui
import time
import cv2
import numpy as np
from mss import mss
import threading

class ImageSearcher:
    def __init__(self):
        self.is_running = True

    def while_search_image_from_monitor(self, monitor_no, image_path):
        self.is_running = True
        monitor = self.get_monitor_info(monitor_no)

        while self.is_running:
            found = self.search_image(monitor, image_path)
            if found:
                x, y = found
                click_x, click_y = x + 20, y + 20
                pyautogui.click(click_x, click_y)
                print(f"Found. X: {x} | Y: {y}")
                self.is_running = False
            else:
                time.sleep(3)

    def search_image_from_monitor(self, monitor_no, image_path):
        monitor = self.get_monitor_info(monitor_no)
        return self.search_image(monitor, image_path) is not None

    def get_monitor_info(self, monitor_no):
        with mss() as sct:
            monitors = sct.monitors
            if monitor_no < 1 or monitor_no > len(monitors):
                raise ValueError("Invalid monitor number")
            return monitors[monitor_no]

    def search_image(self, monitor, image_path):
        with mss() as sct:
            screenshot = np.array(sct.grab(monitor))
        
        template = cv2.imread(image_path, cv2.IMREAD_COLOR)
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        if max_val >= 0.8:  # Threshold for match confidence
            return (max_loc[0] + monitor['left'], max_loc[1] + monitor['top'])
        return None

    def stop_search(self):
        self.is_running = False

def main():
    searcher = ImageSearcher()

    # Example usage:
    # Uncomment and modify these lines as needed
    
    # Run continuous search on monitor 2
    # thread = threading.Thread(target=searcher.while_search_image_from_monitor, args=(2, "finished.png"))
    # thread.start()

    # Search once on monitor 3
    # result = searcher.search_image_from_monitor(3, "finished.png")
    # print(f"Image found: {result}")

    # To stop the continuous search:
    # searcher.stop_search()

if __name__ == "__main__":
    main()