import config
from config import cfg
import util
import cutRecord
import hotkey
import emailNotify

import time
from datetime import datetime, timedelta
import os
import subprocess
from typing import Optional, cast
import cv2
# from cv2.typing import MatLike
import numpy as np
import win32gui
import win32process
import win32api
import win32con
import pywintypes
import psutil
from PIL import ImageGrab
import threading
from pathlib import Path
import copy
import logging
from logging.handlers import RotatingFileHandler

if config.BUNDLE_MODE:
    PIC_TEMPLATE = Path(config.BUNDLE_PATH, "src/monitorMatchTemplates")
else:
    PIC_TEMPLATE = Path(config.EXECUTABLE_DIR, "src/monitorMatchTemplates")
MONITOR_PIC = Path(cfg.paths.otto, r"存档/截图/监视")
MONITOR_LOG_PATH = Path(
    cfg.paths.otto,
    rf'辅助程序/OttoLaserCutting/log/{config.LAUNCH_TIME.strftime("%Y%m%d")}.log'
)


pr = util.pr


class Monitor:
    def __init__(self):
        """Initialize the TubePro monitor with default settings and templates.

        Initializes all monitoring parameters and loads necessary templates. This includes:
        - Monitoring state flags (isRunning, enabled)
        - Timing parameters (checkInterval, alertCooldown)
        - Alert tracking (alertCount, lastAlertTimeStamp)
        - Image matching threshold (similarityThreshold)
        - Template images for state detection

        Attributes:
            isRunning (bool): Flag indicating if monitoring is currently active.
            lastAlertTimeStamp (float): Unix timestamp of last detected alert.
            checkInterval (int): Interval in seconds between monitoring checks.
            checkCount (int): Counter for number of monitoring checks performed.
            programNotFoundRetry (int): Seconds to wait before retrying when program not found.
            alertCooldown (int): Minimum seconds required between consecutive alerts.
            alertHaltThreshold (int): Maximum allowed alerts before halting monitoring.
            alertCount (int): Current count of detected alerts.
            similarityThreshold (float): Threshold (0-1) for template matching similarity.
            enabled (bool): Flag indicating if monitoring is enabled (templates loaded).
            template* (Optional[MatLike]): OpenCV image templates for state detection.
            logger (logging.Logger): Configured logger instance for monitoring events.
        """
        self.isRunning = False
        self.lastAlertTimeStamp = 0.0
        self.checkInterval = 5
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
        self.templateCompletion01 = None
        self.templateCompletion02 = None
        self.templateAlert = None
        self.templateAlertForceReturn = None
        self.templateNoAlert = None
        self.logger = cast(logging.Logger, None)
        # self.templateRunning:          Optional[MatLike] = None
        # self.templatePaused:           Optional[MatLike] = None
        # self.templatePausedCuttingHeadTouch:           Optional[MatLike] = None
        # self.templateCompletion01:       Optional[MatLike] = None
        # self.templateCompletion02:       Optional[MatLike] = None
        # self.templateAlert:            Optional[MatLike] = None
        # self.templateAlertForceReturn: Optional[MatLike] = None
        # self.templateNoAlert:          Optional[MatLike] = None
        self._setupLog()
        self._loadTemplates()


    def _setupLog(self):
        """Initialize and configure the rotating log file handler.

        Sets up a rotating log file with the following characteristics:
        - Checks for and resolves log file name collisions
        - Uses UTF-8 encoding
        - Limits file size to 15MB
        - Keeps 3 backup copies
        - Formats log messages with timestamp and log level

        The logger is stored in self.logger and configured to log INFO level messages.
        """
        # Check duplicated log name collision
        logPath = MONITOR_LOG_PATH
        logPath = util.incrementPathIfExist(logPath)
        os.makedirs(logPath.parent, exist_ok=True)

        # Set up looger
        handler = RotatingFileHandler(
            logPath, # type: ignore
            maxBytes=15 * 1024 * 1024,  # 5 MB
            backupCount=3,
            encoding="utf-8",
        )
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s]: %(message)s")
        )

        self.logger = logging.getLogger("tubeProMonitor")
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(handler)
    # }}}

    def _loadTemplates(self) -> None: # {{{
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
            ("templateCompletion01", "completion01.png"),
            ("templateCompletion02", "completion02.png"),
            ("templateAlert", "alert.png"),
            ("templateAlertForceReturn", "alertForceReturn.png"),
            ("templateNoAlert", "noAlert.png"),
        ]
        for attrName, fileName in templates:
            p = Path(PIC_TEMPLATE, fileName)
            if not p.exists():
                pr(f"Cannot find template: {p}")
                self.logger.error(f"Cannot find template: {p}")
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
                    self.logger.error(f"Error loading template image: {e}")
                    self.enabled = False
                    return
        # }}}

    def _startMonitoring(self) -> None:
        """
        Starts the monitoring process in a separate daemon thread if enabled.
        Prints status messages and returns immediately if monitoring is disabled.
        """
        self.isRunning = True
        pr("Monitoring started.")
        self.logger.info("Monitoring started.")
        threading.Thread(target=self._monitor_loop, daemon=True).start()

    def _stopMonitoring(self) -> None:
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
            self.logger.info("Monitoring stopped.")

    def toggleMonitoring(self) -> None:
        """
        Toggles the monitoring state. If monitoring is running, stops it; otherwise starts it.
        """
        if not self.enabled:
            pr("Monitoring is unavailable due to missing or invalid templates.")
            self.logger.info("Monitoring is unavailable due to missing or invalid templates.")
            return
        if self.isRunning:
            self._stopMonitoring()
        else:
            self._startMonitoring()

    def offWorkShutdownChk(self, currentTime: datetime) -> bool:
        """
        Determines if the current time is within off-work hours (21:00 to next day 07:00).
        Logs whether the machine is wroking during work time.

        Args:
            current_time (datetime): The current datetime to check against work hours.
        """
        # Define the off-work period (21:00 to next day 07:00)
        offWorkStart = datetime(currentTime.year, currentTime.month, currentTime.day, 21, 0, 0)
        workStart = datetime(currentTime.year, currentTime.month, currentTime.day, 7, 0, 0) + timedelta(days=1)

        # Adjust for cases where current_time is before midnight
        if currentTime < offWorkStart:
            offWorkStart -= timedelta(days=1)
            workStart -= timedelta(days=1)

        # Check if current_time is within the off-work period
        if offWorkStart <= currentTime <= workStart:
            return True
        else:
            return False

    def _monitor_loop(self) -> None:
        """
        Monitors the TubePro application window and performs actions based on its state.

        This function continuously checks:
        - If TubePro is the foreground window
        - Mouse cursor idle time (brings TubePro to foreground if idle too long)
        - Matches window content against templates to detect states (paused, completion, alerts)
        - Takes appropriate actions for each state (screenshots, notifications, auto-clicks)
        - Handles error conditions and cooldowns

        The loop runs at intervals defined by self.checkInterval and stops when self.isRunning is False.
        """
        cursorPosLast = None
        cursorPosCurrent = None
        cursorIdleCount = 0
        currentTime = datetime.now()
        tubeProTitleCurrent        = ""
        tubeProTitleLastCompletion = ""
        tubeProTitleLastAlert      = ""
        tubeProTitleLastNormal     = ""
        while self.isRunning:
            tubeProTitleCurrent = ""
            time.sleep(self.checkInterval)
            self.checkCount += 1

            pr(f"Monitoring for the {self.checkCount} times...")
            self.logger.info(f"Monitoring for the {self.checkCount} times...")

            hwndTitles = {}
            def winEnumHandler(hwnd, ctx):
                if win32gui.IsWindowVisible(hwnd):
                    windowText = win32gui.GetWindowText(hwnd)
                    if windowText:
                        hwndTitles[hwnd] = windowText
                return True
            win32gui.EnumWindows(winEnumHandler, None)

            foregroundHWND        = win32gui.GetForegroundWindow()
            foregroundProcessId   = win32process.GetWindowThreadProcessId(foregroundHWND)[1]
            if foregroundProcessId <= 0 or psutil.Process(foregroundProcessId).name() != "TubePro.exe":
                self.logger.warning("TubePro isn't the foreground window.")
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
                                    try:
                                        win32gui.SetForegroundWindow(hwnd)
                                        self.logger.info("TubePro has been idle for too long and now it's been brought to the foreground window.")
                                        cursorIdleCount = 0 # reset
                                    except pywintypes.error:
                                        self.logger.error("Failed to bring tubePro window to the front.")

                                    break


                cursorPosLast = cursorPosCurrent
                continue
            else:
                tubeProHWND = foregroundHWND
                tubeProTitleCurrent = win32gui.GetWindowText(tubeProHWND)

            # Skip if the main window isn't in the front
            if not tubeProTitleCurrent.startswith("TubePro"):
                self.logger.info("Skip due to tubePro isn't focus at the main window.")
                continue

            # Capture window content from TubePro
            screenshot = self.captureWindow(-1)
            if screenshot is None:
                continue

            # Convert to OpenCV format
            screenshotCV = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            # Compare with templates
            for stateName, attrName in (
                ("paused",       "templatePaused"),
                ("completion02", "templateCompletion02"),
                ("alert",        "templateAlert"),
                ("noAlert",      "templateNoAlert"),
            ):
                template = getattr(self, attrName)
                matchResult = cv2.matchTemplate(screenshotCV, template, cv2.TM_CCOEFF_NORMED)
                _, maxVal, _, maxLoc = cv2.minMaxLoc(matchResult)
                if maxVal >= self.similarityThreshold:
                    self.logger.info(f"Matched {stateName} with {maxVal * 100:.2f}% similarity.")
                    if stateName == "completion02": # {{{
                        if tubeProTitleCurrent != tubeProTitleLastCompletion and tubeProTitleCurrent == tubeProTitleLastNormal:
                            tubeProTitleLastCompletion = tubeProTitleCurrent

                            pr(f'Cutting session "{tubeProTitleCurrent}" is completed, taking screenshot record.')
                            self.logger.info(f'Cutting session "{tubeProTitleCurrent}" is completed, taking screenshot record.')

                            # `win32api.MessageBox` inside `takeScreenshot()`
                            # is a blocking call—it halts the thread so we need
                            # to make sure it call in a new thread then
                            # complete thread after 5 seconds
                            cutRecord.takeScreenshot(screenshot)

                            # Make records for monitoring
                            os.makedirs(MONITOR_PIC, exist_ok=True)
                            screenshotPath = util.screenshotSave(screenshot, stateName, MONITOR_PIC)
                            self.logger.info(f"Save screenshot at {screenshotPath}")

                            # Send email notification
                            self.logger.info("Sending email...")
                            emailNotify.send(stateName, tubeProTitleCurrent, screenshotPath)
                            self.logger.info("Email sent")

                            # Check off-work hours and shutdown if necessary
                            if self.offWorkShutdownChk(currentTime):
                                self.isRunning = False
                                subprocess.call(["shutdown", "-s"])
                                self.logger.warning("Currently it's off-work hours, shutdown the machine.")
                            else:
                                self.logger.info("Currently it's work time right now, no plan for shuting down the machine.")

                        break
                    # }}}
                    elif stateName == "paused": # {{{
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

                            if (
                                int(currentTime.timestamp() - self.lastAlertTimeStamp)
                                <= self.alertCooldown
                            ) or self.alertCount >= self.alertHaltThreshold:
                                pr(f"Temperarily disable auto-clicking due to {self.alertCount} times fail in {self.alertCooldown}s")
                                self.logger.warning(f"Stop auto-clicking due to {self.alertCount} times failed in {self.alertCooldown}s")
                                if tubeProTitleCurrent == tubeProTitleLastAlert:
                                    break

                                tubeProTitleLastAlert = tubeProTitleCurrent
                                screenshotPath = util.screenshotSave(screenshot, "pauseAndHalt", MONITOR_PIC)
                                emailNotify.send(stateName, tubeProTitleCurrent, screenshotPath)
                                # Check off-work hours and shutdown if necessary
                                if self.offWorkShutdownChk(currentTime):
                                    self.isRunning = False
                                    subprocess.call(["shutdown", "-s"])
                                    self.logger.warning("Currently it's off-work hours, shutdown the machine.")
                            else:
                                if self.offWorkShutdownChk(currentTime):
                                    self.logger.warning("Cutting is paused, auto-click continue.")
                                    screenshotPath = util.screenshotSave(screenshot, "pauseThenContinue", MONITOR_PIC)
                                    emailNotify.send(stateName, tubeProTitleCurrent, screenshotPath)
                                    savedPosition = copy.copy(hotkey.mouse.position)
                                    time.sleep(5)
                                    hotkey.mouse.position = (maxLoc[0] - 60, maxLoc[1] + 90)
                                    hotkey.mouse.press(hotkey.Button.left)
                                    hotkey.mouse.release(hotkey.Button.left)
                                    hotkey.mouse.position = savedPosition
                                else:
                                    self.logger.info("It' work time now, auto-click continue is disabled.")


                        break
                    # }}}
                    elif stateName == "alert": # {{{
                        self.lastAlertTimeStamp = currentTime.timestamp()
                        matchResultAlertForceReturn = cv2.matchTemplate( # type: ignore
                            screenshotCV,
                            self.templateAlertForceReturn,
                            cv2.TM_CCOEFF_NORMED
                        )
                        _, maxValAlertForceReturn, _, _ = cv2.minMaxLoc(
                            matchResultAlertForceReturn
                        )
                        if tubeProTitleCurrent != tubeProTitleLastAlert:
                            tubeProTitleLastAlert = tubeProTitleCurrent
                            if maxValAlertForceReturn >= self.similarityThreshold:
                                stateName = "alertForceReturn"
                                # TODO: cut down the tube
                                pr("Force return is detected.")
                                self.logger.warning("Force return is detected.")
                            else:
                                pr("Alert is detected.")
                                self.logger.warning("Alert is detected.")

                            screenshotPath = util.screenshotSave(screenshot, stateName, MONITOR_PIC)
                            emailNotify.send(stateName, tubeProTitleCurrent, screenshotPath)

                        break
                    # }}}
                    elif stateName == "noAlert": # {{{
                        matchResultRunning = cv2.matchTemplate( # type: ignore
                            screenshotCV,
                            self.templateAlertForceReturn,
                            cv2.TM_CCOEFF_NORMED
                        )
                        _, maxValRunning, _, _ = cv2.minMaxLoc(matchResultRunning)
                        if maxValRunning >= self.similarityThreshold:
                            tubeProTitleLastNormal = tubeProTitleCurrent
                            if self.alertCount and (currentTime.timestamp() - self.lastAlertTimeStamp > self.alertCooldown):
                                self.alertCount = 0
                                self.tubeProTitleLastAlert = ""
                                self.logger.info("Alert cleared. Back to the track")
                            if tubeProTitleLastCompletion:
                                tubeProTitleLastCompletion = ""
                                self.logger.info("Clear latst completion title. Back to the track")
                        break
                    # }}}

    def checkTemplateMatches(self):
        """
        Checks if the current window screenshot matches any of the predefined templates.
        Compares the screenshot against multiple templates using OpenCV's template matching.
        Prints matching scores for each template and collects matches above similarity threshold.
        Returns None if screenshot capture fails.
        """
        screenshot = self.captureWindow(-1)
        if screenshot is None:
            pr("Caputre image failed")
            self.logger.info("Caputre image failed")
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
            ("completion01",               "templateCompletion01"),
            ("completion02",               "templateCompletion02"),
            ("alert",                      "templateAlert"),
            ("alertForceReturn",           "templateAlertForceReturn"),
            ("noAlert",                    "templateNoAlert"),
        ):
            template = getattr(self, attrName)
            matchResult = cv2.matchTemplate(screenshotCVUint8, template, cv2.TM_CCOEFF_NORMED)
            _, maxVal, _, maxLoc = cv2.minMaxLoc(matchResult)
            pr(f"{name}: {maxVal}")
            self.logger.info(f"{name}: {maxVal}")
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


    def captureWindow(self, hwnd):
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
                self.logger.warning(f"Error capturing window: {e}")
                return None
        else:
            return ImageGrab.grab()




monitor: Optional[Monitor] = None
