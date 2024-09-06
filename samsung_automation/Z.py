import sys
import pyautogui
import time
import cv2
import numpy as np
from mss import mss
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal
# pip install PyQt5 pyautogui mss opencv-python

class ImageSearchThread(QThread):
    found_signal = pyqtSignal(int, int)
    not_found_signal = pyqtSignal()

    def __init__(self, monitor, image_path):
        super().__init__()
        self.monitor = monitor
        self.image_path = image_path
        self.is_running = True

    def run(self):
        while self.is_running:
            found = self.search_image()
            if found:
                self.found_signal.emit(found[0], found[1])
                self.is_running = False
            else:
                self.not_found_signal.emit()
                time.sleep(3)

    def search_image(self):
        with mss() as sct:
            screenshot = np.array(sct.grab(self.monitor))
        
        template = cv2.imread(self.image_path, cv2.IMREAD_COLOR)
        if template is None:
            print("Image file not found or invalid")
            return None

        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        if max_val >= 0.8:  # Threshold for match confidence
            return (max_loc[0] + self.monitor['left'], max_loc[1] + self.monitor['top'])
        return None

    def stop(self):
        self.is_running = False

class ImageSearcherGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.search_thread = None

    def init_ui(self):
        self.setWindowTitle('Image Searcher')
        self.setGeometry(100, 100, 300, 200)

        layout = QVBoxLayout()

        # Monitor selection
        monitor_layout = QHBoxLayout()
        monitor_label = QLabel('Select Monitor:')
        self.monitor_combo = QComboBox()
        self.update_monitor_list()
        monitor_layout.addWidget(monitor_label)
        monitor_layout.addWidget(self.monitor_combo)
        layout.addLayout(monitor_layout)

        # Image selection
        self.image_button = QPushButton('Select Image')
        self.image_button.clicked.connect(self.select_image)
        layout.addWidget(self.image_button)

        # Search buttons
        button_layout = QHBoxLayout()
        self.search_once_button = QPushButton('Search Once')
        self.search_once_button.clicked.connect(self.search_once)
        self.search_continuous_button = QPushButton('Search Continuously')
        self.search_continuous_button.clicked.connect(self.search_continuous)
        button_layout.addWidget(self.search_once_button)
        button_layout.addWidget(self.search_continuous_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        self.image_path = None

    def update_monitor_list(self):
        with mss() as sct:
            for i, monitor in enumerate(sct.monitors[1:], 1):  # Skip the first (all monitors combined)
                self.monitor_combo.addItem(f"Monitor {i}")

    def select_image(self):
        self.image_path, _ = QFileDialog.getOpenFileName(self, "Select Image File", "", "Image Files (*.png *.jpg *.bmp)")
        if self.image_path:
            self.image_button.setText(f"Image: {self.image_path.split('/')[-1]}")

    def get_selected_monitor(self):
        monitor_index = self.monitor_combo.currentIndex() + 1  # +1 because mss monitor list is 1-indexed
        with mss() as sct:
            return sct.monitors[monitor_index]

    def search_once(self):
        if not self.image_path:
            QMessageBox.warning(self, "Warning", "Please select an image first.")
            return

        monitor = self.get_selected_monitor()
        thread = ImageSearchThread(monitor, self.image_path)
        thread.found_signal.connect(self.on_image_found)
        thread.not_found_signal.connect(self.on_image_not_found)
        thread.start()

    def search_continuous(self):
        if not self.image_path:
            QMessageBox.warning(self, "Warning", "Please select an image first.")
            return

        if self.search_thread and self.search_thread.isRunning():
            self.search_thread.stop()
            self.search_thread.wait()
            self.search_continuous_button.setText('Search Continuously')
        else:
            monitor = self.get_selected_monitor()
            self.search_thread = ImageSearchThread(monitor, self.image_path)
            self.search_thread.found_signal.connect(self.on_image_found)
            self.search_thread.not_found_signal.connect(self.on_image_not_found)
            self.search_thread.start()
            self.search_continuous_button.setText('Stop Search')

    def on_image_found(self, x, y):
        QMessageBox.information(self, "Found", f"Image found at X: {x}, Y: {y}")
        click_x, click_y = x + 20, y + 20
        pyautogui.click(click_x, click_y)
        if self.search_thread:
            self.search_thread.stop()
            self.search_continuous_button.setText('Search Continuously')

    def on_image_not_found(self):
        print("Image not found, retrying...")

def main():
    app = QApplication(sys.argv)
    ex = ImageSearcherGUI()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()