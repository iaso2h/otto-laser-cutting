import util
from config import cfg
from console import print
import keySet
import style

import chardet
import os
import re
import datetime
from typing import Tuple, Optional
from collections import Counter
from pathlib import Path
from striprtf.striprtf import rtf_to_text
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Protection
from pprint import pprint

LASER_PROFILE_PATH = Path(cfg.paths.otto, r"存档/耗时计算.xlsx")
TUBEPRO_LOG_PATH   = Path(cfg.paths.otto, r"存档/切割机日志")
laserFileOpenPat = re.compile(r"^\((.+?)\)打开文件：(.+)$")
segmentFirstPat  = re.compile(r"^\((.+?)\)总零件数:(\d+), 当前零件序号:1$")
segmentPat         = re.compile(r".*总零件数:(\d+), 当前零件序号:\d+$")
scheduelTotalPat   = re.compile(r".*零件切割计划数目\d+.*$")
scheduelLoopEndPat = re.compile(r".*已切割零件数目\d+.*$")
loopStartPat       = re.compile(r".*开始加工, 循环计数：\d+.*$")


def getEncoding(filePath) -> str:
    # Create a magic object
    """
    Detects the encoding of a file using chardet.

    Args:
        filePath (str): Path to the file to analyze.

    Returns:
        str: Detected encoding as a string (e.g. 'utf-8'), or empty string if detection fails.
    """
    with open(filePath, "rb") as f:
        # Detect the encoding
        rawData = f.read()
        result = chardet.detect(rawData)
        if not result:
            return ""
        if not result["encoding"]:
            return ""
        else:
            return result["encoding"]


def fillWorkbook(ws: Worksheet, parsedResult: dict, sortChk: bool):
    """
    Fills an Excel worksheet with laser cutting file statistics from parsed data.

    Args:
        ws (Worksheet): OpenPyXL Worksheet object to populate with data.
        parsedResult (dict): Dictionary containing parsed laser file information with keys:
            - loop: List of loop intervals
            - loopIntervalCounter: Counter object of interval frequencies
            - loopIntervalUpdated: Dictionary of last update timestamps per interval
            - workpieceCount: Number of workpieces per file
        sortChk (bool): Whether to sort the results alphabetically by filename.

    Populates worksheet with:
        - Headers in row 1 with formatted columns
        - File statistics grouped by laser filename
        - Calculated fields for material/time consumption
        - Cell protection with password '456'
    """
    ws[f"A{1}"].value = "排样文件"
    ws.column_dimensions["A"].width = 35
    ws[f"B{1}"].value = "循环耗时"
    ws.column_dimensions["B"].width = 12
    ws[f"C{1}"].value = "循环统计"
    ws.column_dimensions["C"].width = 12
    ws[f"D{1}"].value = "最后统计日期"
    ws.column_dimensions["D"].width = 22
    ws[f"E{1}"].value = "工件目标数"
    ws.column_dimensions["E"].width = 14
    ws[f"F{1}"].value = "工件已加工数"
    ws.column_dimensions["F"].width = 17
    ws[f"G{1}"].value = "预计消耗长料"
    ws.column_dimensions["G"].width = 17
    ws[f"H{1}"].value = "预计消耗时长"
    ws.column_dimensions["H"].width = 17
    ws[f"I{1}"].value = "预计完成时间"
    ws.column_dimensions["I"].width = 22
    for col in range(1, 10):
        ws.cell(row=1, column=col).style     = "Headline 1"
        ws.cell(row=1, column=col).alignment = style.alCenter

    if sortChk:
        items = sorted(parsedResult.items())
    else:
        items = parsedResult.items()
    for laserFileName, laserFileInfo in items:
        if len(laserFileInfo["loop"]) < 1:
            continue

        mostCommon = laserFileInfo["loopIntervalCounter"].most_common(5)
        laserFileStartRow = ws.max_row
        skipRowCount = 0
        headlineBorderSet = False
        for intervalIdx, common in enumerate(mostCommon):
            currentRow = intervalIdx + laserFileStartRow + 1 - skipRowCount
            interval      = common[0]
            intervalCount = common[1]

            if intervalIdx == len(mostCommon) - 1:

                # Merge laser filename cells under column A
                if interval == "0":
                    endRow = currentRow - 1
                else:
                    endRow = currentRow
                # Check merge necessity of merging cells
                if endRow > laserFileStartRow + 1:
                    ws.merge_cells(
                        start_row    = laserFileStartRow + 1,
                        end_row      = endRow,
                        start_column = 1,
                        end_column   = 1
                    )
                    ws.cell(row=laserFileStartRow+1,column=1).alignment = style.alCenterWrap

            # Don't fill in sheet when interval between two loop is 0
            if interval == "0":
                skipRowCount += 1
                continue

            ws.cell(row=currentRow, column=2).value = int(float(interval))
            ws.cell(row=currentRow, column=2).number_format = '0"秒"'
            ws.cell(row=currentRow, column=3).value = intervalCount
            ws.cell(row=currentRow, column=3).number_format = '0"次"'
            ws.cell(row=currentRow, column=4).value = laserFileInfo["loopIntervalUpdated"][interval]
            ws.cell(row=currentRow, column=5).value = 100
            ws.cell(row=currentRow, column=5).font = style.font["orangeBold"]
            ws.cell(row=currentRow, column=5).number_format = '0"支"'
            ws.cell(row=currentRow, column=5).protection = Protection(locked=False)
            ws.cell(row=currentRow, column=6).value = 0
            ws.cell(row=currentRow, column=6).font = style.font["greenBold"]
            ws.cell(row=currentRow, column=6).number_format = '0"支"'
            ws.cell(row=currentRow, column=6).protection = Protection(locked=False)
            ws.cell(row=currentRow, column=7).value = f'=(E{currentRow}-F{currentRow})/{laserFileInfo["workpieceCount"]}'
            ws.cell(row=currentRow, column=7).number_format = '0"支"'
            ws.cell(row=currentRow, column=8).value = f'=(B{currentRow}+1)/{laserFileInfo["workpieceCount"]}*(E{currentRow}-F{currentRow})/86400'
            ws.cell(row=currentRow, column=8).number_format = "[h]时mm分ss秒"
            ws.cell(row=currentRow, column=9).value = f'=NOW() + H{currentRow}'
            ws.cell(row=currentRow, column=9).number_format = "yyyy-m-d h:mm:ss"

            if not headlineBorderSet:
                # Add top border
                headlineBorderSet = True
                ws.cell(row=currentRow, column=1).value = laserFileName
                ws[f"A{currentRow}"].border = style.borderMedium
                ws[f"B{currentRow}"].border = style.borderMedium
                ws[f"C{currentRow}"].border = style.borderMedium
                ws[f"D{currentRow}"].border = style.borderMedium
                ws[f"E{currentRow}"].border = style.borderMedium
                ws[f"F{currentRow}"].border = style.borderMedium
                ws[f"G{currentRow}"].border = style.borderMedium
                ws[f"H{currentRow}"].border = style.borderMedium
                ws[f"I{currentRow}"].border = style.borderMedium

            ws.protection.sheet = True
            ws.protection.password = '456'
            ws.protection.enable()


def parse(
    rtfFile: Path,
    wb: Workbook,
    accumulationMode: bool,
    parsedResult: Optional[dict] = None,
) -> dict:
    """
    Parses an RTF file containing laser cutting records and organizes the data into a structured format.

    Args:
        rtfFile: Path to the RTF file to parse
        wb: Workbook object for Excel output (optional, used in non-accumulation mode)
        accumulationMode: If True, accumulates results without writing to Excel
        parsedResult: Optional dictionary to accumulate results across multiple files

    Returns:
        Dictionary containing:
        - "workbook": Modified Workbook object (None in accumulation mode)
        - "parsedResult": Dictionary of parsed data with structure:
            {
                "laserFileName": {
                    "open": [(lineIdx, timestamp)],
                    "loop": [(lineIdx, timestamp, interval)],
                    "loopIntervalUpdated": {interval: timestamp},
                    "loopIntervalCounter": Counter(intervals),
                    "workpieceCount": int
                }
            }

    The function processes RTF content to extract:
    1. Laser file open events with timestamps
    2. Loop segments with timestamps and intervals
    3. Workpiece counts
    4. Statistics on loop intervals
    """
    laserFileLastOpen = ""
    loopLastTime = None
    if parsedResult is None:
        parsedResult = {}

    now = datetime.datetime.now()
    with open(rtfFile, "r", encoding=getEncoding(str(rtfFile))) as f:
        content = rtf_to_text(f.read())
        lines = content.split("\n")

    laserFileFullPath = ""
    for lineIdx, l in enumerate(lines):
        openMatch      = laserFileOpenPat.match(l)
        loopStartMatch = segmentFirstPat.match(l)
        if openMatch:
            laserFileFullPath = openMatch.group(2)
            laserFileName = laserFileFullPath.replace("D:\\欧拓图纸\\切割文件\\", "")
            laserFileName = util.diametartSymbolUnify(laserFileName)
            laserFileName = laserFileName.replace(".zx", ".zzx")
            laserFileName = laserFileName.replace("  ", " ")
            laserFileName = laserFileName.strip()
            laserFileLastOpen = laserFileName
            if laserFileName not in parsedResult:
                parsedResult[laserFileName] = {
                    "open": [],
                    "loop": [],
                    "loopIntervalUpdated": {},
                    "loopIntervalCounter": Counter(),
                    "workpieceCount": 0
                }
                loopLastTime = None
            parsedResult[laserFileName]["open"].append(( lineIdx, openMatch.group(1) ))

        if loopStartMatch:
            timeStamp = loopStartMatch.group(1)
            timeLoop  = datetime.datetime.strptime(f"{now.year}/{timeStamp}", "%Y/%m/%d %H:%M:%S")
            if not loopLastTime:
                loopInterval = 0
            else:
                loopInterval = (timeLoop - loopLastTime).total_seconds()

            # Add addiontional time window in accumulation mode
            if accumulationMode:
                if loopInterval:
                    loopInterval += 15

            loopLastTime = timeLoop

            parsedResult[laserFileLastOpen]["loop"].append(( lineIdx, timeStamp, loopInterval))
            parsedResult[laserFileLastOpen]["loopIntervalUpdated"][f"{loopInterval}"] = loopLastTime
            parsedResult[laserFileLastOpen]["loopIntervalCounter"][f"{loopInterval}"] += 1
            # Get maximun workpiece count
            if int(loopStartMatch.group(2)) > parsedResult[laserFileLastOpen]["workpieceCount"]:
                parsedResult[laserFileLastOpen]["workpieceCount"] = int(loopStartMatch.group(2))

    if not parsedResult:
        print(f"No laser file records parsed from rtf file {str(rtfFile)}")
        return {
            "workbook":     None,
            "parsedResult": None
        }

    if accumulationMode:
        return {
                "workbook":     None,
                "parsedResult": parsedResult
                }
    else:
        if wb.active.title == "Sheet": # type: ignore
            ws = wb.active # type: ignore
            ws.title = rtfFile.stem # type: ignore
        else:
            ws = wb.create_sheet(rtfFile.stem, 0)
        fillWorkbook(ws, parsedResult, True) # type: ignore
        return {
                "workbook":     wb,
                "parsedResult": parsedResult
                }


def parseAllLog():
    """
    Parses all RTF log files in the specified directory (excluding files with '精简' in their names),
    processes them using the parse() function, and saves the combined results to an Excel workbook.

    Returns:
        None: Outputs the result to a file rather than returning a value.
    """
    wb = Workbook()
    for f in Path(TUBEPRO_LOG_PATH).iterdir():
        if f.suffix != ".rtf" or "精简" in f.stem:
            continue

        wb = parse(
            rtfFile=f,
            wb=wb,
            accumulationMode=False
                )["workbook"] # type: ignore
    util.saveWorkbook(wb, LASER_PROFILE_PATH, True) # type: ignore


def parseAccuLog():
    """
    Parses accumulated laser cutting logs from RTF files within the last 60 days.
    Processes files in TUBEPRO_LOG_PATH (excluding '精简' files), extracts data into a workbook,
    and saves results to LASER_PROFILE_PATH. Skips hidden column F in output.

    Returns:
        None: Prints message if no logs found, otherwise saves processed data to file.
    """
    wb = Workbook()
    parsedResult = None
    now = datetime.datetime.now()
    timeDeltaLiteral = 60
    timeDelta = datetime.timedelta(days=timeDeltaLiteral)

    for f in Path(TUBEPRO_LOG_PATH).iterdir():
        if f.suffix != ".rtf" or "精简" in f.stem:
            continue

        logTime = datetime.datetime.fromtimestamp(f.stat().st_ctime)
        if now - logTime <= timeDelta:
            parsedResult = parse(
                rtfFile=f,
                wb=wb,
                accumulationMode=True,
                parsedResult=parsedResult
                    )["parsedResult"] # type: ignore
    if not parsedResult:
        return print("No parsed accumulated result")

    fillWorkbook(wb.active, parsedResult, True) # type: ignore
    wb.active.column_dimensions['F'].hidden = True #type: ignore
    util.saveWorkbook(wb, LASER_PROFILE_PATH, True) # type: ignore


def parsePeriodLog():
    """
    Parses laser cutting log files within a specified time period based on modifier keys.
    Handles different parsing modes:
    - Ctrl+Shift+Alt: Parse accumulated logs (calls parseAccuLog)
    - Ctrl: Open laser profile directly
    - Shift: Parse logs from last 7 days
    - Alt: Parse all logs (calls parseAllLog)
    - No modifier: Parse logs from last 1 day
    Automatically expands time window (up to 3 attempts) if no logs found.
    Saves parsed data to laser profile if logs were processed.
    """
    if "ctrl" in keySet.keys and "shift" in keySet.keys and "alt" in keySet.keys:
        return parseAccuLog()
    elif "ctrl" in keySet.keys:
        return os.startfile(LASER_PROFILE_PATH)
    elif "shift" in keySet.keys:
        timeDeltaLiteral = 7
    elif "alt" in keySet.keys:
        return parseAllLog()
    else:
        timeDeltaLiteral = 1

    wb = Workbook()
    now = datetime.datetime.now()
    parsedPeriodCount = 0
    for loopCount in range(3):
        if parsedPeriodCount > 0:
            break
        else:
            # Increase the time delta window to if no parsed files
            timeDeltaLiteral = timeDeltaLiteral * (7 ** loopCount)
        timeDelta = datetime.timedelta(days=timeDeltaLiteral)

        for f in Path(TUBEPRO_LOG_PATH).iterdir():
            if f.suffix != ".rtf" or "精简" in f.stem:
                continue

            logTime = datetime.datetime.fromtimestamp(f.stat().st_ctime)
            if now - logTime <= timeDelta:
                wb = parse(
                    rtfFile=f,
                    wb=wb,
                    accumulationMode=False
                    )["workbook"] # type: ignore
                parsedPeriodCount += 1

    if parsedPeriodCount:
        util.saveWorkbook(wb, LASER_PROFILE_PATH, True) # type: ignore


def rtfSimplify():
    """
    Processes RTF log files in TUBEPRO_LOG_PATH based on modifier keys:
    - Ctrl: Opens the log directory
    - Shift: Processes files from last 7 days
    - Alt: Processes files from last year
    - No modifier: Processes files from last day

    For each matching RTF file:
    1. Filters content using regex patterns (laserFileOpenPat, segmentPat, etc.)
    2. Creates a simplified version with '精简' prefix in filename
    3. Outputs processed files with relevant log lines

    Handles cases where no files are found by expanding time window exponentially.
    """
    if "ctrl" in keySet.keys:
        return os.startfile(TUBEPRO_LOG_PATH)
    elif "shift" in keySet.keys:
        timeDeltaLiteral = 7
    elif "alt" in keySet.keys:
        timeDeltaLiteral = 360
    else:
        timeDeltaLiteral = 1

    now = datetime.datetime.now()
    parsedPeriodCount = 0
    for loopCount in range(3):
        if parsedPeriodCount > 0:
            break
        else:
            # Increase the time delta window to if no parsed files
            timeDeltaLiteral = timeDeltaLiteral * (7 ** loopCount)
        timeDelta = datetime.timedelta(days=timeDeltaLiteral)

        for f in TUBEPRO_LOG_PATH.iterdir():
            if f.suffix != ".rtf" or "精简" in f.stem:
                continue

            rtfTime = datetime.datetime.fromtimestamp(f.stat().st_ctime)
            if now - rtfTime <= timeDelta:
                with open(f, "r", encoding=getEncoding(str(f))) as f1:
                    content = rtf_to_text(f1.read())
                    lines = content.split("\n")
                refineLines = []
                for line in lines:
                    m1 = laserFileOpenPat.match(line)
                    m2 = segmentPat.match(line)
                    m3 = scheduelTotalPat.match(line)
                    m4 = scheduelLoopEndPat.match(line)
                    m5 = loopStartPat.match(line)
                    if not any((m1, m2, m3, m4, m5)):
                        continue
                    else:
                        refineLines.append(line + "\n")
                targetPath = Path(
                    f.parent,
                    "精简" + f.stem + f.suffix
                )
                with open(targetPath, mode="w", encoding="utf-8") as f2:
                    for line in refineLines:
                        f2.write(line)
                print("导出日志: ", str(targetPath))

    print("rtf日志精简完成")
