import config
import console
import os
import shutil
import datetime
import win32api, win32con
import re
from pathlib import Path
from openpyxl import Workbook
from typing import List


print = console.print


def getTimeStamp() -> str:
    now = datetime.datetime.now()
    return str(now.strftime("%H:%M:%S"))
    # return str(now.strftime(f"%Y/{now.month}/%d %H:%M:%S"))


def saveWorkbook(wb: Workbook, dstPath: Path | None = None, openAfterSaveChk=False) -> Path: # {{{
    os.makedirs(config.LOCAL_EXPORT_DIR, exist_ok=True)
    timeStr = str(datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f"))

    if dstPath and (os.getlogin() == "OT03" or config.DEV_MODE):
        # Create backup first
        if dstPath.exists():
            backupPath = Path(
                config.LOCAL_EXPORT_DIR,
                dstPath.stem + "_backup_" + timeStr + ".xlsx"
            )
            shutil.copy2(dstPath, backupPath)

        try:
            wb.save(str(dstPath))
            print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{dstPath}")
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
                    config.LOCAL_EXPORT_DIR,
                    dstPath.stem + "_fallback_" + timeStr + ".xlsx")
                wb.save(str(fallbackExcelPath))
                print(f"\n[{getTimeStamp()}]:[bold green]Saving fallback Excel file at: [/bold green][bright_black]{fallbackExcelPath}")
                return fallbackExcelPath

    else:
        newExcelPath = Path(
            config.LOCAL_EXPORT_DIR,
            timeStr + ".xlsx")
        wb.save(str(newExcelPath))
        print(f"\n[{getTimeStamp()}]:[bold green]Saving new Excel file at: [/bold green][bright_black]{newExcelPath}")
        if openAfterSaveChk:
            os.startfile(newExcelPath)
        return newExcelPath



def strStandarize(old: Path) -> Path:
    if old.is_file():
        new = str(old)
        # old = old.replace("∅", "∅")
        new = new.replace("Ø", "∅")
        new = new.replace("Φ", "∅")
        new = new.replace("φ", "∅")
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
