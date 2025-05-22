import config
import util

import time
import cv2
import numpy as np
import win32gui
import win32process
import psutil
from PIL import ImageGrab
import threading
from pathlib import Path
from typing import Optional

print = util.print
logger = util.monitorLogger
monitor = None

class Monitor:
    def __init__(self):
        self.isRunning = False
        self.template = None
        self.templateHeight = 0
        self.templateWidth = 0
        self.lastAlertTime = 0
        self.checkInterval = 1
        self.checkCount = 0
        self.programNotFoundRetry = 60
        self.alertCooldown = 60
        self.alertShutdonwThreshold = 3
        self.alertShutdonwCount = 0
        self.similarityThreshold = 0.9
        self.enabled = True

        normalTemplatePath = Path(config.PIC_TEMPLATE, "runningNormal.png") # type: ignore
        if not normalTemplatePath.exists():
            self.enabled = False
        else:
            self.loadTemplate(str(normalTemplatePath))

    def loadTemplate(self, templatePathStr: str):
        try:
            self.template = cv2.imdecode(np.fromfile(templatePathStr, dtype=np.uint8), cv2.IMREAD_COLOR)
            if self.template is None:
                raise FileNotFoundError(f"Template image not found at {templatePathStr}")
            self.templateHeight, self.templateWidth = self.template.shape[:2]
            return True
        except Exception as e:
            print(f"Error loading template image: {e}")
            return False


    def startMonitoring(self):
        self.isRunning = True
        print("Monitoring started.")
        threading.Thread(target=self._monitor_loop, daemon=True).start()


    def stopMonitoring(self):
        self.isRunning = False
        print("Monitoring stopped.")

    def toggleMonitoring(self):
        if self.isRunning:
            self.stopMonitoring()
        else:
            self.startMonitoring()

    def _monitor_loop(self):
        while self.isRunning:
            self.checkCount += 1
            logger.info(f"Monitoring for the {self.checkCount} times...")
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
                if pName.lower() == "tubepro.exe":
                    targetHWND = hwnd
                    break

            if targetHWND == -1:
                print(f"Tubepro browser not found. Retry in {self.programNotFoundRetry}s")
                time.sleep(self.programNotFoundRetry)
                continue

            # Capture window content
            screenshot = captureWindow(targetHWND)
            if screenshot is None:
                print(f"Caputre image failed")
                time.sleep(self.checkInterval)
                continue

            # Convert to OpenCV format
            screenshotCV = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            # Compare with template
            matchResult = cv2.matchTemplate(screenshotCV, self.template, cv2.TM_CCOEFF_NORMED) # type: ignore
            _, maxVal, _, maxLoc = cv2.minMaxLoc(matchResult)

            # Check for match
            currentTime = time.time()
            if maxVal < self.similarityThreshold:
                print("Match failed.")
                self.alertShutdonwCount += 1
                self.lastAlertTime = currentTime
                if (currentTime - self.lastAlertTime < self.alertCooldown) and self.alertShutdonwCount >= self.alertShutdonwThreshold:
                    self.isRunning = False
                    self.alertShutdonwCount = 0
                    print(f"Stop monitoring due to {self.alertShutdonwCount} times fail in {self.alertCooldown}s")
                # print(f"ALERT! Match found: {maxVal * 100:.2f}% similarity")
                #
                # # Visualize match (optional)
                # topLeft = maxLoc
                # bottomRight = (topLeft[0] + self.templateWidth, topLeft[1] + self.templateHeight)
                # cv2.rectangle(screenshotCV, topLeft, bottomRight, (0, 255, 0), 2)
                # cv2.imshow('Match Found', screenshotCV)
                # cv2.waitKey(10000)
                # cv2.destroyAllWindows()
            else:
                print("Match succeeded.")
                if (currentTime - self.lastAlertTime >= self.alertCooldown):
                    self.alertShutdonwCount = 0
                logger.info("Everything is fine")

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
    if monitor.enabled:
        monitor.startMonitoring()
    else:
        return

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        monitor.stopMonitoring()
        print("\nMonitoring stopped.")
    finally:
        cv2.destroyAllWindows()
