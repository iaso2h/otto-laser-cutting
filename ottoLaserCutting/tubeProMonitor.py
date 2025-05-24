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
from cv2.typing import MatLike

print = util.print
logger = util.monitorLogger
monitor = None

class Monitor:
    def __init__(self):
        self.isRunning = False
        self.templateHeight = 0
        self.templateWidth = 0
        self.lastAlertTime = 0
        self.checkInterval = 3
        self.checkCount = 0
        self.programNotFoundRetry = 60
        self.alertCooldown = 60
        self.alertShutdonwThreshold = 3
        self.alertShutdonwCount = 0
        self.similarityThreshold = 0.9
        self.enabled = True
        self.templateRunning:          Optional[MatLike] = None
        self.templatePaused:           Optional[MatLike] = None
        self.templateFinished01:       Optional[MatLike] = None
        self.templateFinished02:       Optional[MatLike] = None
        self.templateAlert:            Optional[MatLike] = None
        self.templateAlertForceReturn: Optional[MatLike] = None
        self.templateNoAlert:          Optional[MatLike] = None


    def loadTemplates(self) -> None:
        """Set up different templates"""
        # Check existences of all templates
        templates = [
            ("templateRunning",          "running.png"),
            ("templatePaused",           "paused.png"),
            ("templateFinished01",       "finished01.png"),
            ("templateFinished02",       "finished02.png"),
            ("templateAlert",            "alert.png"),
            ("templateAlertForceReturn", "alertForceReturn.png"),
            ("templateNoAlert",          "noAlert.png")
        ]
        for attrName, fileName in templates:
            p = Path(config.PIC_TEMPLATE, fileName)  # type: ignore
            if not p.exists():
                print(f"Cannot find template: {p}")
                self.enabled = False
                return
            else:
                try:
                    template = cv2.imdecode(np.fromfile(p, dtype=np.uint8), cv2.IMREAD_COLOR)
                    if template is None:
                        raise FileNotFoundError(f"Template image not found at {p}")
                    setattr(self, attrName, template)
                except Exception as e:
                    print(f"Error loading template image: {e}")
                    self.enabled = False
                    return



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

    @staticmethod
    def getTubeProHWND() -> int:
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

        return targetHWND


    def _monitor_loop(self):
        while self.isRunning:
            time.sleep(self.checkInterval)
            self.checkCount += 1

            logger.info(f"\n\nMonitoring for the {self.checkCount} times...")

            foregroundHWND = win32gui.GetForegroundWindow()
            foregroundProcessId = win32process.GetWindowThreadProcessId(foregroundHWND)[1]
            foregroundProcessName = psutil.Process(foregroundProcessId).name()
            if foregroundProcessName != "TubePro.exe":
                logger.info(f"TubePro isn't the foreground window.")
                continue
            else:
                tubeProHWND = foregroundHWND
            # Find TubePro window
            # tubeProHWND = self.getTubeProHWND()
            # if tubeProHWND == -1:
            #     logger.info(f"Tubepro not found. Retry in {self.programNotFoundRetry}s")
            #     continue

            # Capture window content from TubePro
            screenshot = captureWindow(tubeProHWND)
            if screenshot is None:
                logger.info(f"Caputre image failed")
                continue

            # Convert to OpenCV format
            screenshotCV = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            # Compare with templates
            for name, attrName in (
                ("running",          "templateRunning"),
                ("paused",           "templatePaused"),
                ("finished01",       "templateFinished01"),
                ("finished02",       "templateFinished02"),
                ("alert",            "templateAlert"),
                ("alertForceReturn", "templateAlertForceReturn"),
                ("noAlert",          "templateNoAlert"),
            ):
                template = getattr(self, attrName)
                matchResult = cv2.matchTemplate(screenshotCV, template, cv2.TM_CCOEFF_NORMED)
                _, maxVal, _, maxLoc = cv2.minMaxLoc(matchResult)
                if maxVal >= self.similarityThreshold:
                    logger.info(f"→{name}: {maxVal}]←")
                else:
                    logger.info(f"{name}: {maxVal}")
            currentTime = time.time()

            # if maxVal < self.similarityThreshold:
            #     print(f"Match failed at similarity {maxVal}.")
            #     self.alertShutdonwCount += 1
            #     self.lastAlertTime = currentTime
            #     if (currentTime - self.lastAlertTime < self.alertCooldown) and self.alertShutdonwCount >= self.alertShutdonwThreshold:
            #         self.isRunning = False
            #         self.alertShutdonwCount = 0
            #         print(f"Stop monitoring due to {self.alertShutdonwCount} times fail in {self.alertCooldown}s")
            #         return
            #     # print(f"ALERT! Match found: {maxVal * 100:.2f}% similarity")
            # else:
            #     print("Match succeeded.")
            #     if (currentTime - self.lastAlertTime >= self.alertCooldown):
            #         self.alertShutdonwCount = 0
            #     logger.info("Everything is fine")


    def checkTemplateMatches(self):
        screenshot = captureWindow(-1)
        if screenshot is None:
            print(f"Caputre image failed")
            return

        # Convert to OpenCV format
        screenshotCV      = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        # Make sure it's CV_8U. Credit: https://stackoverflow.com/a/33184916/10273260
        # screenshotCVUint8 = screenshotCV.astype(np.uint8)
        screenshotCVUint8 = screenshotCV

        # Compare with template
        print("--------------")
        matchResults = []
        for name, attrName in (
            ("running",          "templateRunning"),
            ("paused",           "templatePaused"),
            ("finished01",       "templateFinished01"),
            ("finished02",       "templateFinished02"),
            ("alert",            "templateAlert"),
            ("alertForceReturn", "templateAlertForceReturn"),
            ("noAlert",          "templateNoAlert"),
        ):
            template = getattr(self, attrName)
            matchResult = cv2.matchTemplate(screenshotCVUint8, template, cv2.TM_CCOEFF_NORMED)
            _, maxVal, _, maxLoc = cv2.minMaxLoc(matchResult)
            print(f"{name}: {maxVal}")
            if maxVal >= self.similarityThreshold:
                templateWidth, templateHeight = template.shape[:2]
                matchResults.append((maxVal, maxLoc, templateWidth, templateHeight))
        # Highlight matched area
        # for m in matchResults:
        #     maxVal, topLeft, templateWidth, templateHeight = m
        #     bottomRight = (topLeft[0] + templateWidth, topLeft[1] + templateHeight)
        #     cv2.rectangle(screenshotCV, topLeft, bottomRight, (0, 255, 0), 2)
        #     cv2.imshow('Match Found', screenshotCV)
        #     cv2.waitKey(10000)
        #     cv2.destroyAllWindows()



def captureWindow(hwnd):
    """Capture window content using Pillow."""
    if hwnd != -1:
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            return ImageGrab.grab(bbox=(left, top, right, bottom))
        except Exception as e:
            print(f"Error capturing window: {e}")
            return None
    else:
        return ImageGrab.grab()


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
