import time
import cv2
import numpy as np
import win32gui
import win32process
import psutil
from PIL import ImageGrab
import threading

class Monitor:
    def __init__(self):
        self.isRunning = False
        self.template = None
        self.templateHeight = 0
        self.templateWidth = 0
        self.lastAlertTime = 0
        self.checkInterval = 1
        self.alertCooldown = 5
        self.similarityThreshold = 0.8

    def loadTemplate(self, templatePath):
        try:
            self.template = cv2.imread(templatePath, cv2.IMREAD_COLOR)
            if self.template is None:
                raise FileNotFoundError(f"Template image not found at {templatePath}")
            self.templateHeight, self.templateWidth = self.template.shape[:2]
            return True
        except Exception as e:
            print(f"Error loading template image: {e}")
            return False

    def startMonitoring(self):
        if not self.template:
            print("Template not loaded. Call loadTemplate() first.")
            return

        self.isRunning = True
        print("Monitoring started.")
        threading.Thread(target=self._monitor_loop, daemon=True).start()

    def stopMonitoring(self):
        self.isRunning = False
        print("Monitoring stopped.")

    def _monitor_loop(self):
        while self.isRunning:
            # Find target window
            hwndTitle = {}
            def winEnumHandler(hwnd, ctx):
                if win32gui.IsWindowVisible(hwnd):
                    windowText = win32gui.GetWindowText(hwnd)
                    if windowText:
                        hwndTitle[hwnd] = windowText
                return True

            win32gui.EnumWindows(winEnumHandler, None)

            targetHWND = -1
            for hwnd, title in hwndTitle.items():
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                pName = psutil.Process(pid).name()
                if pName.lower() == "vivaldi.exe":
                    targetHWND = hwnd
                    break

            if targetHWND == -1:
                print("Vivaldi browser not found.")
                time.sleep(self.checkInterval)
                continue

            # Capture window content
            screenshot = captureWindow(targetHWND)
            if screenshot is None:
                time.sleep(self.checkInterval)
                continue

            # Convert to OpenCV format
            screenshotCV = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            # Compare with template
            matchResult = cv2.matchTemplate(screenshotCV, self.template, cv2.TM_CCOEFF_NORMED)
            _, maxVal, _, maxLoc = cv2.minMaxLoc(matchResult)

            # Check for match
            currentTime = time.time()
            if maxVal >= self.similarityThreshold and currentTime - self.lastAlertTime > self.alertCooldown:
                print(f"ALERT! Match found: {maxVal * 100:.2f}% similarity")
                self.lastAlertTime = currentTime

                # Visualize match (optional)
                topLeft = maxLoc
                bottomRight = (topLeft[0] + self.templateWidth, topLeft[1] + self.templateHeight)
                cv2.rectangle(screenshotCV, topLeft, bottomRight, (0, 255, 0), 2)
                cv2.imshow('Match Found', screenshotCV)
                cv2.waitKey(10000)
                cv2.destroyAllWindows()
            else:
                print("No match found.")

            time.sleep(self.checkInterval)

def captureWindow(hwnd):
    """Capture window content using Pillow."""
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        return ImageGrab.grab(bbox=(left, top, right, bottom))
    except Exception as e:
        print(f"Error capturing window: {e}")
        return None

def main():
    monitor = Monitor()
    monitor.loadTemplate(r"")
    monitor.startMonitoring()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        monitor.stopMonitoring()
        print("\nMonitoring stopped.")
    finally:
        cv2.destroyAllWindows()
