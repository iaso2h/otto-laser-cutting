import config
from config import cfg
import util
import cutRecord
import hotkey
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
pr = util.pr
logger = util.monitorLogger
monitor = None


class Monitor:
    def __init__(self):
        """
        Initializes the TubeProMonitor instance with default values for monitoring state and templates.
        Attributes:
            isRunning (bool): Indicates if monitoring is active.
            lastAlertTimeStamp (float): Timestamp of last alert.
            checkInterval (int): Seconds between monitoring checks.
            checkCount (int): Number of checks performed.
            programNotFoundRetry (int): Seconds to wait before retrying after program not found.
            alertCooldown (int): Minimum seconds between alerts.
            alertHaltThreshold (int): Max alerts before halting monitoring.
            alertCount (int): Current alert count.
            similarityThreshold (float): Image similarity threshold for detection.
            enabled (bool): Whether monitoring is enabled.
            template* (Optional[MatLike]): Image templates for various monitoring states.
        """
        self.isRunning = False
        self.lastAlertTimeStamp = 0.0
        self.checkIntervalNormal = 3
        self.checkIntervalLong  = 180
        self.checkInterval = self.checkIntervalNormal
        self.checkCount = 0
        self.programNotFoundRetry = 60
        self.alertCooldown = 60
        self.alertHaltThreshold = 3
        self.alertCount = 0
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
        # self.templatePausedCuttingHeadTouch:           Optional[MatLike] = None
        # self.templateFinished01:       Optional[MatLike] = None
        # self.templateFinished02:       Optional[MatLike] = None
        # self.templateAlert:            Optional[MatLike] = None
        # self.templateAlertForceReturn: Optional[MatLike] = None
        # self.templateNoAlert:          Optional[MatLike] = None

    def loadTemplates(self) -> None:
        """
        Loads and validates OpenCV template images for monitoring tube processing states.

        This method:
        1. Defines a list of required template images with their corresponding attribute names.
        2. Checks if each template file exists in the specified directory (PIC_TEMPLATE).
        3. Attempts to load valid images using OpenCV, storing them as instance attributes.
        4. Sets 'enabled' flag to False and returns early if any template is missing or invalid.

        The loaded templates will be available as instance attributes with the specified names.
        Raises no exceptions but prints error messages for missing/invalid templates.
        """
        # Check existences of all templates
        templates = [
            ("templateRunning", "running.png"),
            ("templatePaused", "paused.png"),
            ("templatePausedCuttingHeadTouch", "pausedWithCuttingHeadTouch.png"),
            ("templateFinished01", "finished01.png"),
            ("templateFinished02", "finished02.png"),
            ("templateAlert", "alert.png"),
            ("templateAlertForceReturn", "alertForceReturn.png"),
            ("templateNoAlert", "noAlert.png"),
        ]
        for attrName, fileName in templates:
            p = Path(PIC_TEMPLATE, fileName)
            if not p.exists():
                pr(f"Cannot find template: {p}")
                logger.error(f"Cannot find template: {p}")
                self.enabled = False
                return
            else:
                try:
                    template = cv2.imdecode(np.fromfile(p, dtype=np.uint8), cv2.IMREAD_COLOR)
                    if template is None:
                        raise FileNotFoundError(f"Template image not found at {p}")
                    setattr(self, attrName, template)
                except Exception as e:
                    pr(f"Error loading template image: {e}")
                    logger.error(f"Error loading template image: {e}")
                    self.enabled = False
                    return

    def startMonitoring(self) -> None:
        """
        Starts the monitoring process in a separate daemon thread if enabled.
        Prints status messages and returns immediately if monitoring is disabled.
        """
        self.isRunning = True
        pr("Monitoring started.")
        logger.info("Monitoring started.")
        threading.Thread(target=self._monitor_loop, daemon=True).start()

    def stopMonitoring(self) -> None:
        """Stops the monitoring process if it's currently running.

        Displays a confirmation dialog asking whether to stop monitoring.
        Only works when monitoring is enabled (valid templates exist).

        Returns:
            None: Prints status messages but doesn't return any value.
        """
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
            pr("Monitoring stopped.")
            logger.info("Monitoring stopped.")

    def toggleMonitoring(self) -> None:
        """
        Toggles the monitoring state. If monitoring is running, stops it; otherwise starts it.
        """
        if not self.enabled:
            pr("Monitoring is unavailable due to missing or invalid templates.")
            logger.info("Monitoring is unavailable due to missing or invalid templates.")
            return
        if self.isRunning:
            self.stopMonitoring()
        else:
            self.startMonitoring()

    def offWorkShutdownChk(self, currentTime: datetime):
        """
        Shuts down the machine during off-work hours (21:00 to next day 07:00).
        If current time falls within this period, sets isRunning flag to False and initiates system shutdown.

        Args:
            currentTime (datetime): The current datetime to check against work hours.

        Logs:
            Info message indicating whether shutdown was triggered or not.
        """
        midNight = datetime(
            currentTime.year, currentTime.month, currentTime.day, 0, 0, 0
        )
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
            pr("Currently it's off-work hours, shutdown the machine.")
            logger.warning("Currently it's off-work hours, shutdown the machine.")

    def _monitor_loop(self) -> None:
        """
        Monitors the TubePro application window and performs actions based on its state.

        This function continuously checks:
        - If TubePro is the foreground window
        - Mouse cursor idle time (brings TubePro to foreground if idle too long)
        - Matches window content against templates to detect states (paused, finished, alerts)
        - Takes appropriate actions for each state (screenshots, notifications, auto-clicks)
        - Handles error conditions and cooldowns

        The loop runs at intervals defined by self.checkInterval and stops when self.isRunning is False.
        """
        cursorPosLast = None
        cursorPosCurrent = None
        cursorIdleCount = 0
        completionIdleCount = 0
        currentTime = datetime.now()
        tubeProTitleCurrent        = ""
        tubeProTitleLastCompletion = ""
        while self.isRunning:
            tubeProTitleCurrent = ""
            time.sleep(self.checkInterval)
            self.checkCount += 1

            pr(f"Monitoring for the {self.checkCount} times...")
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
                pr("TubePro isn't the foreground window.", gui=False)
                logger.warning("TubePro isn't the foreground window.")
                cursorPosCurrent = hotkey.mouse.position

                if cursorPosLast:
                    # Increase the count when cursor stays in the same
                    # position as the last iteration does. Reset the count to
                    # 0 when the ordinates alter
                    if cursorPosCurrent == cursorPosLast:
                        cursorIdleCount += 1
                    else:
                        cursorIdleCount = 0

                    # Set to foreground if TubePro is actually running and being idle for over 1 minutes
                    if cursorIdleCount >= 60 // self.checkInterval:
                        for hwnd, title in hwndTitles.items():
                            if title.startswith("TubePro"):
                                _, pId = win32process.GetWindowThreadProcessId(hwnd)
                                pName = psutil.Process(pId).name()
                                if pName == "TubePro.exe":
                                    tubeProTitleCurrent = title
                                    if win32gui.IsIconic(hwnd):
                                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                                    win32gui.SetForegroundWindow(hwnd)
                                    pr("TubePro has been idle for too long and now it's been brought to the foreground window")
                                    logger.info("TubePro has been idle for too long and now it's been brought to the foreground window")
                                    cursorIdleCount = 0 # reset
                                break


                cursorPosLast = cursorPosCurrent
                continue
            else:
                tubeProHWND = foregroundHWND
                tubeProTitleCurrent = win32gui.GetWindowText(tubeProHWND)

            # Capture window content from TubePro
            screenshot = captureWindow(-1)
            if screenshot is None:
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
                    pr(f"Matched {name} with {maxVal * 100:.2f}% similarity.", gui=False)
                    logger.info(f"Matched {name} with {maxVal * 100:.2f}% similarity.")
                    if name == "finished02": # {{{
                        if tubeProTitleCurrent != tubeProTitleLastCompletion:
                            tubeProTitleLastCompletion = tubeProTitleCurrent

                            pr(f'Cutting session "{tubeProTitleCurrent}" is completed, taking screenshot record.')
                            logger.info(f'Cutting session "{tubeProTitleCurrent}" is completed, taking screenshot record.')

                            # `win32api.MessageBox` inside `takeScreenshot()`
                            # is a blocking call—it halts the thread so we need
                            # to make sure it call in a new thread then
                            # complete thread after 5 seconds
                            curRecordThread = threading.Thread(target=lambda: cutRecord.takeScreenshot(screenshot))
                            curRecordThread.start()
                            messageBoxHwnd = cutRecord.findMessageBoxWindow()
                            if messageBoxHwnd:
                                time.sleep(5)
                                ctypes.windll.user32.PostMessageW(messageBoxHwnd, win32con.WM_CLOSE, 0, 0)

                            curRecordThread.join() # Ensure the thread completes

                            # Make records for monitoring
                            os.makedirs(MONITOR_PIC, exist_ok=True)
                            screenshotPath = util.screenshotSave(screenshot, name, MONITOR_PIC)

                            # Send email notification
                            emailNotify.send(name, tubeProTitleCurrent, screenshotPath)


                            # Check off-work hours and shutdown if necessary
                            self.offWorkShutdownChk(currentTime)
                        else:
                            completionIdleCount += 1
                            if completionIdleCount >= 60 and self.checkInterval == self.checkIntervalNormal:
                                self.checkIntervalNormal = self.checkIntervalLong

                        break
                    # }}}
                    elif name == "paused": # {{{
                        matchResultPausedCuttingHeadTouch = cv2.matchTemplate( # type: ignore
                            screenshotCV,
                            self.templatePausedCuttingHeadTouch,
                            cv2.TM_CCOEFF_NORMED
                        )
                        _, maxValPausedCuttingHeadTouch, _, _ = cv2.minMaxLoc(
                            matchResultPausedCuttingHeadTouch
                        )
                        if maxValPausedCuttingHeadTouch >= self.similarityThreshold:
                            self.alertCount += 1
                            self.lastAlertTimeStamp = currentTime.timestamp()
                            name = "pauseThenContinue"

                            if (
                                (
                                    int(
                                        currentTime.timestamp()
                                        - self.lastAlertTimeStamp
                                        )
                                    < self.alertCooldown
                                )
                                or self.alertCount
                                >= self.alertHaltThreshold
                            ):
                                pr(f"Stop auto-clicking due to {self.alertCount} times fail in {self.alertCooldown}s")
                                logger.warning(f"Stop auto-clicking due to {self.alertCount} times fail in {self.alertCooldown}s")
                                if self.checkInterval == self.checkIntervalNormal:
                                    self.checkInterval = self.checkIntervalLong
                                util.screenshotSave(screenshot, "pauseAndHalt", MONITOR_PIC)
                                self.offWorkShutdownChk(currentTime)
                            else:
                                pr("Cutting is paused, auto-click continue.")
                                logger.info("Cutting is paused, auto-click continue.")
                                screenshotPath = util.screenshotSave(screenshot, "pauseThenContinue", MONITOR_PIC)
                                emailNotify.send(name, tubeProTitleCurrent, screenshotPath)
                                savedPosition = copy.copy(hotkey.mouse.position)
                                time.sleep(5)
                                hotkey.mouse.position = (maxLoc[0] - 60, maxLoc[1] + 90)
                                hotkey.mouse.press(hotkey.Button.left)
                                hotkey.mouse.release(hotkey.Button.left)
                                hotkey.mouse.position = savedPosition

                        break
                    # }}}
                    elif name == "alert": # {{{
                        self.lastAlertTimeStamp = currentTime.timestamp()
                        matchResultAlertForceReturn = cv2.matchTemplate( # type: ignore
                            screenshotCV,
                            self.templateAlertForceReturn,
                            cv2.TM_CCOEFF_NORMED
                        )
                        _, maxValAlertForceReturn, _, _ = cv2.minMaxLoc(
                            matchResultAlertForceReturn
                        )
                        if maxValAlertForceReturn >= self.similarityThreshold:
                            name = "alertForceReturn"
                            # TODO: cut down the tube
                            pr("Force return is detected.")
                            logger.warning("Force return is detected.")
                            if self.checkInterval == self.checkIntervalNormal:
                                self.checkInterval = self.checkIntervalLong
                            screenshotPath = util.screenshotSave(screenshot, "alertForceReturn", MONITOR_PIC)
                            emailNotify.send(name, tubeProTitleCurrent, screenshotPath)
                        else:
                            pr("Alert is detected.")
                            logger.warning("Alert is detected.")
                            emailNotify.send(name, tubeProTitleCurrent)
                            if self.checkInterval == self.checkIntervalNormal:
                                self.checkInterval = self.checkIntervalLong
                            screenshotPath = util.screenshotSave(screenshot, "alert", MONITOR_PIC)
                            emailNotify.send(name, tubeProTitleCurrent, screenshotPath)

                        break
                    # }}}
                    elif name == "noAlert": # {{{
                        matchResultRunning = cv2.matchTemplate( # type: ignore
                            screenshotCV,
                            self.templateAlertForceReturn,
                            cv2.TM_CCOEFF_NORMED
                        )
                        _, maxValRunning, _, _ = cv2.minMaxLoc(matchResultRunning)
                        if maxValRunning >= self.similarityThreshold:
                            if self.alertCount and (currentTime.timestamp() - self.lastAlertTimeStamp >= self.alertCooldown):
                                self.alertCount = 0
                                pr("Alert cleared. Back to the track")
                                logger.info("Alert cleared. Back to the track")

                            completionIdleCount = 0

                            if self.checkInterval != self.checkIntervalNormal:
                                self.checkInterval = self.checkIntervalNormal

                        break
                    # }}}

    def checkTemplateMatches(self):
        """
        Checks if the current window screenshot matches any of the predefined templates.
        Compares the screenshot against multiple templates using OpenCV's template matching.
        Prints matching scores for each template and collects matches above similarity threshold.
        Returns None if screenshot capture fails.
        """
        screenshot = captureWindow(-1)
        if screenshot is None:
            pr(f"Caputre image failed")
            logger.info(f"Caputre image failed")
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
            pr(f"{name}: {maxVal}")
            logger.info(f"{name}: {maxVal}")
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
    """
    Captures the content of a specified window or the entire screen if no window is specified.

    Args:
        hwnd: The handle to the window to capture. If -1, captures the entire screen.

    Returns:
        PIL.Image.Image: The captured image as a Pillow Image object, or None if an error occurs.

    Raises:
        Prints any exception that occurs during capture but does not raise it.
    """
    if hwnd != -1:
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            return ImageGrab.grab(bbox=(left, top, right, bottom))
        except Exception as e:
            pr(f"Error capturing window: {e}")
            logger.warning(f"Error capturing window: {e}")
            return None
    else:
        return ImageGrab.grab()
