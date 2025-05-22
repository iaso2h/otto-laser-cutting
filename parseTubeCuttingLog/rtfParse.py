import util
import config
import console
import style
import keySet

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

laserFileOpenPat = re.compile(r"^\((.+?)\)打开文件：(.+)$")
segmentFirstPat  = re.compile(r"^\((.+?)\)总零件数:(\d+), 当前零件序号:1$")
segmentPat         = re.compile(r".*总零件数:(\d+), 当前零件序号:\d+$")
scheduelTotalPat   = re.compile(r".*零件切割计划数目\d+.*$")
scheduelLoopEndPat = re.compile(r".*已切割零件数目\d+.*$")
loopStartPat       = re.compile(r".*开始加工, 循环计数：\d+.*$")

def getWorkbook():
    if config.CUT_RECORD_PATH.exists():
        return load_workbook(str(config.LASER_PORFILING_PATH))
    else:
        return Workbook()
print = console.print


def getEncoding(filePath) -> str:
    # Create a magic object
    with open(filePath, "rb") as f:
        # Detect the encoding
        rawData = f.read()
        result = chardet.detect(rawData)
        if not result["encoding"]:
            return ""
        else:
            return result["encoding"]


def fillWorkbook(ws: Worksheet, parsedResult: dict, sortChk: bool):
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
    parsedResult: Optional[dict] = None
) -> dict:
    """
    Args:
        parsedResult: If None, a new dictionary will be created.
                      Provide a dict if you want to accumulate results.
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
            ws = wb.active # type: Worksheet
            ws.title = rtfFile.stem
        else:
            ws = wb.create_sheet(rtfFile.stem, 0)
        fillWorkbook(ws, parsedResult, True)
        return {
                "workbook":     wb,
                "parsedResult": parsedResult
                }


def parseAllLog():
    wb = Workbook()
    for f in Path(config.TUBEPRO_LOG_PATH).iterdir():
        if f.suffix != ".rtf" or "精简" in f.stem:
            continue

        wb = parse(
            rtfFile=f,
            wb=wb,
            accumulationMode=False
                )["workbook"] # type: ignore
    util.saveWorkbook(wb, config.LASER_PROFILE_PATH, True) # type: ignore


def parseAccuLog():
    wb = Workbook()
    parsedResult = None
    now = datetime.datetime.now()
    timeDeltaLiteral = 60
    timeDelta = datetime.timedelta(days=timeDeltaLiteral)

    for f in Path(config.TUBEPRO_LOG_PATH).iterdir():
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
    util.saveWorkbook(wb, config.LASER_PROFILE_PATH, True) # type: ignore


def parsePeriodLog():
    if "ctrl" in keySet.keys and "shift" in keySet.keys and "alt" in keySet.keys:
        return parseAccuLog()
    elif "ctrl" in keySet.keys:
        return os.startfile(config.LASER_PROFILE_PATH)
    elif "shift" in keySet.keys:
        timeDeltaLiteral = 7
    elif "alt" in keySet.keys:
        return parseAllLog()
    else:
        timeDeltaLiteral = 1
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

        for f in Path(config.TUBEPRO_LOG_PATH).iterdir():
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
        util.saveWorkbook(wb, config.LASER_PROFILE_PATH, True) # type: ignore


def rtfSimplify():
    for p in config.TUBEPRO_LOG_PATH.iterdir():
        if p.suffix != ".rtf" or "精简" in p.stem:
            continue

        with open(p, "r", encoding=getEncoding(str(p))) as f1:
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
            p.parent,
            "精简" + p.stem + p.suffix
        )
        with open(targetPath, mode="w", encoding="utf-8") as f2:
            for line in refineLines:
                f2.write(line)
        print("导出日志: ", targetPath)

