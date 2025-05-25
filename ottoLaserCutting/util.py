import config
from config import cfg

import os
import shutil
import datetime
import win32api, win32con
import re
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from openpyxl import Workbook
from typing import List
from PIL.Image import Image

MONITOR_LOG_PATH = Path(cfg.paths.otto, r"存档/切割机监视.log")

# Logging set up
handler = RotatingFileHandler(
    MONITOR_LOG_PATH, # type: ignore
    maxBytes=5 * 1024 * 1024,  # 5 MB
    backupCount=3,
    encoding="utf-8",
)
handler.setFormatter(
    logging.Formatter("%(asctime)s [%(levelname)s]: %(message)s")
)

monitorLogger = logging.getLogger("tubeProMonitor")
monitorLogger.setLevel(logging.INFO)
monitorLogger.addHandler(handler)


def getTimeStamp() -> str:
    now = datetime.datetime.now()
    return str(now.strftime("%H:%M:%S"))
    # return str(now.strftime(f"%Y/{now.month}/%d %H:%M:%S"))


def saveWorkbook(wb: Workbook, dstPath: Path | None = None, openAfterSaveChk=False) -> Path: # {{{
    fallbackExportDir = Path(config.EXECUTABLE_DIR, "export")
    timeStr = str(datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f"))
    os.makedirs(fallbackExportDir, exist_ok=True)

    if dstPath and (os.getlogin() == "OT03" or config.DEV_MODE):
        # Create backup first
        if dstPath.exists():
            backupPath = Path(
                fallbackExportDir,
                dstPath.stem + "_backup_" + timeStr + ".xlsx"
            )
            shutil.copy2(dstPath, backupPath)

        try:
            wb.save(str(dstPath))
            print(f"\n[{getTimeStamp()}]:Saving Excel file at: {dstPath}")
            if openAfterSaveChk:
                os.startfile(dstPath)
            return dstPath
        except PermissionError:
            if win32con.IDRETRY == win32api.MessageBox(
                None,
                f"是否要重新写入该路径？\n\"{str(dstPath)}\"",
                "写入权限不足",
                4096 + 5 + 32
                ):
                #   MB_SYSTEMMODAL==4096
                ##  Button Styles:
                ### 0:OK  --  1:OK|Cancel -- 2:Abort|Retry|Ignore -- 3:Yes|No|Cancel -- 4:Yes|No -- 5:Retry|No -- 6:Cancel|Try Again|Continue
                ##  To also change icon, add these values to previous number
                ### 16 Stop-sign  ### 32 Question-mark  ### 48 Exclamation-point  ### 64 Information-sign ('i' in a circle)
                return saveWorkbook(wb, dstPath, openAfterSaveChk)
            else:
                fallbackExcelPath = Path(
                    fallbackExportDir,
                    dstPath.stem + "_fallback_" + timeStr + ".xlsx")
                wb.save(str(fallbackExcelPath))
                print(f"\n[{getTimeStamp()}]:Saving fallback Excel file at: {fallbackExcelPath}")
                return fallbackExcelPath

    else:
        newExcelPath = Path(
            fallbackExportDir,
            timeStr + ".xlsx")
        wb.save(str(newExcelPath))
        print(f"\n[{getTimeStamp()}]:Saving new Excel file at: {newExcelPath}")
        if openAfterSaveChk:
            os.startfile(newExcelPath)
        return newExcelPath



def strStandarize(old: Path) -> Path:
    if old.is_file():
        new = str(old)
        new = diametartSymbolUnify(new)
        new = new.replace("_T1_", "_T1.0_")
        new = new.replace("xT1x", "xT1.0x")
        new = re.sub(r"\s{2,}", " ", new)
        newPath = Path(new)

        if str(old) != str(newPath) and newPath.exists():
            if old.stat().st_mtime > newPath.stat().st_mtime:
                os.remove(newPath)
            else:
                os.remove(old)
                return old

        try:
            os.rename(old, new)
            return Path(new)
        except PermissionError as e:
            print(str(e))
            return old

    else:
        return old


def getAllLaserFiles() -> List[Path]: # {{{
    laserFilePaths = []

    if not config.LASER_FILE_DIR_PATH.exists():
        return laserFilePaths

    for p in config.LASER_FILE_DIR_PATH.iterdir():
        p = strStandarize(p)
        if p.is_file() and "demo" not in p.stem.lower():
            laserFilePaths.append(p)

    return laserFilePaths # }}}

def diametartSymbolUnify(input: str) -> str:
    # input = input.replace("∅", "∅")
    input = input.replace("Ø", "∅")
    input = input.replace("Φ", "∅")
    input = input.replace("φ", "∅")
    return input


def screenshotSave(screenshot: Image, namePrefix:str, dstDirPath: Path) -> Path:
    os.makedirs(dstDirPath, exist_ok=True)
    datetimeNow = datetime.datetime.now()
    screenshotPath = Path(
            dstDirPath,
            f'{namePrefix} {datetimeNow.strftime("%Y-%m-%d %H%M%S")}.png'
            )
    screenshot.save(screenshotPath)
    return screenshotPath

