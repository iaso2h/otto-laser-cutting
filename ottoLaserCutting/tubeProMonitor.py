import config
from config import cfg
import util
import cutRecord
import hotkey
from console import print
import emailNotify

import time
from datetime import datetime, timedelta
import os
import ctypes
import subprocess
from typing import Optional
import cv2
# from cv2.typing import MatLike
import numpy as np
import win32gui, win32process, win32api, win32con
import psutil
from PIL import ImageGrab
import threading
from pathlib import Path
import copy

if config.BUNDLE_MODE:
    PIC_TEMPLATE = Path(config.BUNDLE_PATH, "src/monitorMatchTemplates")
else:
    PIC_TEMPLATE = Path(config.EXECUTABLE_DIR, "src/monitorMatchTemplates")
MONITOR_PIC = Path(cfg.paths.otto, r"存档/截图/监视")
logger = util.monitorLogger
monitor = None

class Monitor:
    def __init__(self):
        self.isRunning = False
        self.templateHeight = 0
        self.templateWidth = 0
        self.lastAlertTimeStamp = 0.0
        self.checkInterval = 3
        self.checkCount = 0
        self.programNotFoundRetry = 60
        self.alertCooldown = 60
        self.alertShutdonwThreshold = 3
        self.alertShutdownCount = 0
        self.similarityThreshold = 0.9
        self.enabled = True
        self.templateRunning = None
        self.templatePaused = None
        self.templatePausedCuttingHeadTouch = None
        self.templateFinished01 = None
        self.templateFinished02 = None
        self.templateAlert = None
        self.templateAlertForceReturn = None
        self.templateNoAlert = None
        # self.templateRunning:          Optional[MatLike] = None
        # self.templatePaused:           Optional[MatLike] = None
        # self.templateFinished01:       Optional[MatLike] = None
        # self.templateFinished02:       Optional[MatLike] = None
        # self.templateAlert:            Optional[MatLike] = None
        # self.templateAlertForceReturn: Optional[MatLike] = None
        # self.templateNoAlert:          Optional[MatLike] = None


    def loadTemplates(self) -> None:
        """Set up different Opencv templates"""
        # Check existences of all templates
        templates = [
            ("templateRunning",                "running.png"),
            ("templatePaused",                 "paused.png"),
            ("templatePausedCuttingHeadTouch", "pausedWithCuttingHeadTouch.png"),
            ("templateFinished01",             "finished01.png"),
            ("templateFinished02",             "finished02.png"),
            ("templateAlert",                  "alert.png"),
            ("templateAlertForceReturn",       "alertForceReturn.png"),
            ("templateNoAlert",                "noAlert.png")
        ]
        for attrName, fileName in templates:
            p = Path(PIC_TEMPLATE, fileName)
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



    def startMonitoring(self) -> None:
        self.isRunning = True
        print("Monitoring started.")
        threading.Thread(target=self._monitor_loop, daemon=True).start()


    def stopMonitoring(self) -> None:
        self.isRunning = False
        ans = win32api.MessageBox(
                    None,
                    "监视已开启, 是否停止监视?",
                    "监视询问",
                    4096 + 64 + 4
                )
                #   MB_SYSTEMMODAL==4096
                ##  Button Styles:
                ### 0:OK  --  1:OK|Cancel -- 2:Abort|Retry|Ignore -- 3:Yes|No|Cancel -- 4:Yes|No -- 5:Retry|No -- 6:Cancel|Try Again|Continue
                ##  To also change icon, add these values to previous number
                ### 16 Stop-sign  ### 32 Question-mark  ### 48 Exclamation-point  ### 64 Information-sign ('i' in a circle)
        if ans == win32con.IDYES:
            print("Monitoring stopped.")


    def toggleMonitoring(self) -> None:
        if self.isRunning:
            self.stopMonitoring()
        else:
            self.startMonitoring()


    def shutdownOffWorkTime(self, currentTime: datetime):
        midNight = datetime(currentTime.year, currentTime.month, currentTime.day, 0, 0, 0)
        midNight += timedelta(days=1)
        if currentTime < midNight:
            timeGetOffWork = datetime(currentTime.year, currentTime.month, currentTime.day, 21, 0, 0)
            timeGoToWork   = datetime(currentTime.year, currentTime.month, currentTime.day, 7, 0, 0)
            timeGoToWork += timedelta(days=1)
        else:
            timeGetOffWork = datetime(currentTime.year, currentTime.month, currentTime.day, 21, 0, 0)
            timeGetOffWork -= timedelta(days=1)
            timeGoToWork   = datetime(currentTime.year, currentTime.month, currentTime.day, 7, 0, 0)
        if timeGetOffWork <= currentTime <= timeGoToWork:
            self.isRunning = False
            subprocess.call(["shutdown", "-s"])



    def _monitor_loop(self) -> None:
        cursorPosLast = None
        cursorPosCurrent = None
        cursorIdleCount = 0
        currentTime = datetime.now()
        tubeProTitleCurrent        = ""
        tubeProTitleLastCompletion = ""
        while self.isRunning:
            tubeProTitleCurrent = ""
            time.sleep(self.checkInterval)
            self.checkCount += 1

            logger.info("")
            logger.info(f"Monitoring for the {self.checkCount} times...")

            hwndTitles = {}
            def winEnumHandler(hwnd, ctx):
                if win32gui.IsWindowVisible(hwnd):
                    windowText = win32gui.GetWindowText(hwnd)
                    if windowText:
                        hwndTitles[hwnd] = windowText
                return True
            win32gui.EnumWindows(winEnumHandler, None)

            foregroundHWND = win32gui.GetForegroundWindow()
            foregroundProcessId = win32process.GetWindowThreadProcessId(foregroundHWND)[1]
            foregroundProcessName = psutil.Process(foregroundProcessId).name()
            if foregroundProcessName != "TubePro.exe":
                logger.info(f"TubePro isn't the foreground window.")
                cursorPosCurrent = hotkey.mouse.position

                if cursorPosLast:
                    if cursorPosCurrent == cursorPosLast:
                        cursorIdleCount =+ 1

                    # Set to foreground if TubePro is actually running and being idle for over 1 minutes
                    if cursorIdleCount >= 60 // self.checkInterval:
                        for hwnd, title in hwndTitles.items():
                            if title.startswith("TubePro"):
                                _, pId = win32process.GetWindowThreadProcessId(hwnd)
                                pName = psutil.Process(pId).name()
                                if pName == "TubePro.exe":
                                    tubeProTitleCurrent = title
                                    win32gui.ShowWindow(hwnd, 5)
                                    win32gui.SetForegroundWindow(hwnd)
                                    logger.info(f"TubePro has been idle for too long and now it's being brought to the foreground window")
                            elif title == cutRecord.MESSAGEBOX_TITLE:
                                ctypes.windll.user32.PostMessageW(hwnd, win32con.WM_CLOSE, 0, 0)


                        cursorIdleCount = 0 # reset

                cursorPosLast = cursorPosCurrent
                continue
            else:
                tubeProHWND = foregroundHWND
                tubeProTitleCurrent = win32gui.GetWindowText(tubeProHWND)

            # Capture window content from TubePro
            screenshot = captureWindow(-1)
            if screenshot is None:
                logger.info(f"Caputre image failed")
                continue

            # Convert to OpenCV format
            screenshotCV = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            # Compare with templates
            for name, attrName in (
                ("paused",     "templatePaused"),
                ("finished02", "templateFinished02"),
                ("alert",      "templateAlert"),
                ("noAlert",    "templateNoAlert"),
            ):
                template = getattr(self, attrName)
                matchResult = cv2.matchTemplate(screenshotCV, template, cv2.TM_CCOEFF_NORMED)
                _, maxVal, _, maxLoc = cv2.minMaxLoc(matchResult)
                if maxVal >= self.similarityThreshold:
                    logger.info(f"Matched {name} with {maxVal * 100:.2f}% similarity.")
                    if name == "finished02":
                        if tubeProTitleCurrent != tubeProTitleLastCompletion:
                            tubeProTitleLastCompletion = tubeProTitleCurrent
                            cutRecord.takeScreenshot(screenshot)
                            logger.info("Cutting session is completed, stop monitoring.")
                            os.makedirs(MONITOR_PIC, exist_ok=True)
                            util.screenshotSave(screenshot, name, MONITOR_PIC)
                            self.shutdownOffWorkTime(currentTime)
                    elif name == "paused":
                        matchResultPausedCuttingHeadTouch = cv2.matchTemplate(
                            screenshotCV,
                            self.templatePausedCuttingHeadTouch,
                            cv2.TM_CCOEFF_NORMED
                        )
                        _, maxValPausedCuttingHeadTouch, _, _ = cv2.minMaxLoc(
                            matchResultPausedCuttingHeadTouch
                        )
                        if maxValPausedCuttingHeadTouch >= self.similarityThreshold:
                            self.alertShutdownCount += 1
                            self.lastAlertTimeStamp = currentTime.timestamp()
                            if (
                                (
                                    int(
                                        currentTime.timestamp()
                                        - self.lastAlertTimeStamp
                                        )
                                    < self.alertCooldown
                                )
                                and self.alertShutdownCount
                                >= self.alertShutdonwThreshold
                            ):
                                print(f"Stop monitoring due to {self.alertShutdownCount} times fail in {self.alertCooldown}s")
                                self.isRunning = False
                                self.alertShutdownCount = 0
                                util.screenshotSave(screenshot, "pauseAndHalt", MONITOR_PIC)
                                self.shutdownOffWorkTime(currentTime)
                                break
                            else:
                                print(f"Cutting is paused, auto-click continue.")
                                util.screenshotSave(screenshot, "pauseThenContinue", MONITOR_PIC)
                                savedPosition = copy.copy(hotkey.mouse.position)
                                hotkey.mouse.position = (maxLoc[0] - 60, maxLoc[1] + 90)
                                hotkey.mouse.press(hotkey.Button.left)
                                hotkey.mouse.release(hotkey.Button.left)
                                hotkey.mouse.position = savedPosition
                    elif name == "alert":
                        matchResultAlertForceReturn = cv2.matchTemplate(
                            screenshotCV,
                            self.templateAlertForceReturn,
                            cv2.TM_CCOEFF_NORMED
                        )
                        _, maxValAlertForceReturn, _, _ = cv2.minMaxLoc(
                            matchResultAlertForceReturn
                        )
                        if maxValAlertForceReturn >= self.similarityThreshold:
                            # TODO: cut down the tube
                            logger.info("Force return is detected, stop monitoring.")
                            emailNotify.send("Force return is detected, stop monitoring.")
                            self.isRunning = False
                            util.screenshotSave(screenshot, "alertForceReturn", MONITOR_PIC)
                            break
                        else:
                            logger.info("Alert is detected, stop monitoring.")
                            self.isRunning = False
                            util.screenshotSave(screenshot, "alert", MONITOR_PIC)
                            break
                    elif name == "noAlert":
                        matchResultRunning = cv2.matchTemplate(
                            screenshotCV,
                            self.templateAlertForceReturn,
                            cv2.TM_CCOEFF_NORMED
                        )
                        _, maxValAlertRunning, _, _ = cv2.minMaxLoc(matchResultRunning)
                        if (
                            maxValAlertRunning >= self.similarityThreshold
                            and self.alertShutdonwCount
                            and (currentTime.timestamp() - self.lastAlertTimeStamp >= self.alertCooldown)
                        ):
                            self.alertShutdonwCount = 0
                            logger.info("Clear alert count reseted. Back to the track")

                        break


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
        matchResults = []
        for name, attrName in (
            ("running",                    "templateRunning"),
            ("paused",                     "templatePaused"),
            ("pausedWithCuttingHeadTouch", "templatePausedCuttingHeadTouch"),
            ("finished01",                 "templateFinished01"),
            ("finished02",                 "templateFinished02"),
            ("alert",                      "templateAlert"),
            ("alertForceReturn",           "templateAlertForceReturn"),
            ("noAlert",                    "templateNoAlert"),
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
