import config
from config import cfg

import os
import shutil
import datetime
import win32api, win32con
import re
import dearpygui.dearpygui as dpg
from pprint import pprint
from pathlib import Path
from openpyxl import Workbook
from typing import Optional
from typing import List
from PIL.Image import Image


logFlow = []


def pr(*args, gui: bool=True):
    """
    Custom print function that redirects output to a Dear PyGui log window.
    Maintains a global log buffer (logFlow) and updates the GUI log display.

    Args:
        *args: Strings to be printed (will be joined with newlines).
        **kwargs: Unused, maintained for print() signature compatibility,
    """

    global logFlow
    newMessage = "\n".join(args)
    logFlow.append(newMessage)
    message = "\n".join(logFlow)
    if gui:
        dpg.set_value("log", value=message)
    if not config.BUNDLE_MODE:
        pprint(newMessage)


def getTimeStamp(dt: Optional[datetime.datetime]=None) -> str:
    """
    Returns current time in 'HH:MM:SS' format as a string.

    Returns:
        str: Formatted time string showing hours, minutes and seconds.
    """
    if not dt:
        dt = datetime.datetime.now()
    return dt.strftime("%H:%M:%S")
    # return str(now.strftime(f"%Y/{now.month}/%d %H:%M:%S"))


def saveWorkbook(
    wb: Workbook, dstPath: Path | None = None, openAfterSaveChk=False
) -> Path:  # {{{
    """
    Saves a Workbook object to specified path with fallback options.

    Args:
        wb: Workbook object to be saved
        dstPath: Optional destination path for the workbook. If None or permission issues occur,
                 falls back to default export directory.
        openAfterSaveChk: If True, opens the saved file after saving

    Returns:
        Path: The actual path where the workbook was saved

    Behavior:
        - Creates backup if file exists
        - Handles permission errors with retry prompt
        - Falls back to export directory if:
            * No dstPath provided
            * Permission error occurs and user chooses not to retry
            * Not running as OT03 user outside DEV_MODE
        - Generates timestamped filenames for fallback files
        - Can optionally open saved file after saving
    """
    fallbackExportDir = Path(config.EXECUTABLE_DIR, "export")
    timeStr = str(datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f"))
    os.makedirs(fallbackExportDir, exist_ok=True)

    if dstPath:
        # Create backup first
        if dstPath.exists():
            backupPath = Path(
                fallbackExportDir,
                dstPath.stem + "_backup_" + timeStr + ".xlsx"
            )
            shutil.copy2(dstPath, backupPath)

        try:
            wb.save(str(dstPath))
            pr(f"\n[{getTimeStamp()}]:Saving Excel file at: {dstPath}")
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
                pr(f"\n[{getTimeStamp()}]:Saving fallback Excel file at: {fallbackExcelPath}")
                return fallbackExcelPath

    else:
        newExcelPath = Path(
            fallbackExportDir,
            timeStr + ".xlsx")
        wb.save(str(newExcelPath))
        pr(f"\n[{getTimeStamp()}]:Saving new Excel file at: {newExcelPath}")
        if openAfterSaveChk:
            os.startfile(newExcelPath)
        return newExcelPath


def strStandarize(srcPath: Path) -> Path:
    """
    Standardizes a file path string by performing the following operations:
    1. Unifies diameter symbols using `diametartSymbolUnify`.
    2. Replaces "_T1_" and "xT1x" with "_T1.0_" and "xT1.0x" respectively.
    3. Collapses multiple spaces into single spaces.
    4. Handles file renaming with timestamp comparison:
       - If destination exists, keeps the newer file.
       - Returns original path if rename fails due to permissions.

    Args:
        oldPath (Path): Original file path to standardize

    Returns:
        Path: Standardized path if successful, original path otherwise
    """
    if srcPath.is_file():
        refinedName = str(srcPath)
        refinedName = diametartSymbolUnify(refinedName)
        refinedName = refinedName.replace("_T1_", "_T1.0_")
        refinedName = refinedName.replace("xT1x", "xT1.0x")
        refinedName = re.sub(r"\s{2,}", " ", refinedName)
        refinedPath = Path(refinedName)

        if str(srcPath) != str(refinedPath) and refinedPath.exists():
            if srcPath.stat().st_mtime > refinedPath.stat().st_mtime:
                os.remove(refinedPath)
                try:
                    os.rename(srcPath, refinedPath)
                    return refinedPath
                except PermissionError as e:
                    pr(str(e))
                    return srcPath
            else:
                os.remove(srcPath)
                return refinedPath
        else:
            try:
                os.rename(srcPath, refinedPath)
                return refinedPath
            except PermissionError as e:
                pr(str(e))
                return srcPath
    else:
        return srcPath


def getAllLaserFiles() -> List[Path]:  # {{{
    """
    Retrieves all laser cutting files from the configured directory, excluding demo files.

    Returns:
        List[Path]: A list of Path objects representing valid laser cutting files.
        Empty list if the directory doesn't exist.
    """
    laserFilePaths = []

    if not config.LASER_FILE_DIR_PATH.exists():
        return laserFilePaths

    for p in config.LASER_FILE_DIR_PATH.iterdir():
        p = strStandarize(p)
        if p.is_file() and "demo" not in p.stem.lower():
            laserFilePaths.append(p)

    return laserFilePaths
# }}}


def diametartSymbolUnify(input: str) -> str:
    # input = input.replace("∅", "∅")
    """
    Unifies different diameter symbol representations to a single standard symbol (∅).

    Args:
        input (str): The string containing diameter symbols to be unified.

    Returns:
        str: The input string with all diameter symbols replaced by the standard ∅ symbol.
    """
    input = input.replace("Ø", "∅")
    input = input.replace("Φ", "∅")
    input = input.replace("φ", "∅")
    return input


def screenshotSave(screenshot: Image, namePrefix: str, dstDirPath: Path) -> Path:
    """
    Saves a screenshot image to the specified directory with a timestamped filename.

    Args:
        screenshot: PIL Image object to be saved
        namePrefix: Prefix string for the filename
        dstDirPath: Destination directory path where the image will be saved

    Returns:
        Path: The full path where the screenshot was saved

    Example:
        >>> img = Image.new('RGB', (100, 100))
        >>> path = screenshotSave(img, 'test', Path('/screenshots'))
        >>> pr(path)  # e.g. /screenshots/test 2023-01-01 120000.png
    """
    os.makedirs(dstDirPath, exist_ok=True)
    datetimeNow = datetime.datetime.now()
    screenshotPath = Path(
        dstDirPath, f'{namePrefix} {datetimeNow.strftime("%Y-%m-%d %H%M%S")}.png'
    )
    screenshot.save(screenshotPath)
    return screenshotPath


fileNameIncreamentPat = re.compile(r"^(.*)\((\d+)\)$")
def incrementPathIfExist(p: Path) -> Path:
    if not p.exists():
        return p

    duplicateCount = 1
    while True:
        match = fileNameIncreamentPat.match(p.stem)
        if match:
            duplicateCount = int(match.group(2))
            duplicateCount += 1
            p = Path(
                    p.parent,
                    fileNameIncreamentPat.sub(
                        rf"\1({duplicateCount})",
                        p.stem
                    ) + p.suffix
            )
            if not p.exists():
                return p
        else:
            duplicateCount += 1
            p = Path(
                    p.parent,
                    p.stem + f"({ duplicateCount })" + p.suffix,
            )
            if not p.exists():
                return p
