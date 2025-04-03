# File: parseTubeProLog
# Author: iaso2h
# Description: Parsing Log files(.rtf) from TubePro and split them into separated files
VERSION     = "0.0.82"
LASTUPDATED = "2025-04-03"

import os
from pathlib import Path

GUI_MODE    = False
SILENT_MODE = False
DEV_MODE    = False
PROGRAM_DIR = Path(os.getcwd())
LOCAL_EXPORT_DIR = Path(PROGRAM_DIR, "export")
RE_LASER_FILES_PAT = r"^(\d{3}[-a-zA-Z]{,3}(\(.+?\))?) (([^_]+? )?[^_]+?) ([^_]{,8})([_*x])?(([^\u4e00-\u9fff]+?).(([RT])?([0-9.]+?))[_*x]L([.0-9]{1,5}))?( \d+?支 \+ (.*?) \d+?支)?( L(\d{4}))?(_X\d{,2})?"
TUBE_DIMENSION_PAT = r"(∅[0-9.]*?)(\*T.*?)?\*(L.*)"


PARENT_DIR_PATH = Path(r"D:\欧拓图纸")
WAREHOUSING_PATH = Path(r"E:\Stock\外协")

LASER_FILE_DIR_PATH  = ""
DISPATCH_FILE_PATH   = ""
SCREENSHOT_DIR_PATH  = ""
CUT_RECORD_PATH      = ""
LASER_PORFILING_PATH = ""
LASER_LOG_PATH     = ""
LASER_PROFILE_PATH = ""
LASER_OCR_FIX_PATH                   = ""
PRODUCT_ID_CATERGORY_CONVENTION_PATH = ""
GUI_GEOMETRY_PATH = ""
WORKPIECE_DICT = ""

def updaPath():
    global LASER_FILE_DIR_PATH
    global DISPATCH_FILE_PATH
    global SCREENSHOT_DIR_PATH
    global CUT_RECORD_PATH
    global LASER_PORFILING_PATH
    global LASER_LOG_PATH
    global LASER_PROFILE_PATH
    global LASER_OCR_FIX_PATH
    global PRODUCT_ID_CATERGORY_CONVENTION_PATH
    global GUI_GEOMETRY_PATH
    global WORKPIECE_DICT

    LASER_FILE_DIR_PATH  = Path(PARENT_DIR_PATH, r"切割文件")
    DISPATCH_FILE_PATH   = Path(PARENT_DIR_PATH, r"派工单（模板+空表）.xlsx")
    SCREENSHOT_DIR_PATH  = Path(PARENT_DIR_PATH, r"存档/截图")
    CUT_RECORD_PATH      = Path(PARENT_DIR_PATH, r"存档/开料记录.xlsx")
    LASER_PORFILING_PATH = Path(PARENT_DIR_PATH, r"存档/开料耗时.xlsx")
    LASER_PROFILE_PATH   = Path(PARENT_DIR_PATH, r"存档/耗时计算.xlsx")
    LASER_LOG_PATH       = Path(PARENT_DIR_PATH, r"存档/切割机日志")
    LASER_OCR_FIX_PATH                   = Path(PARENT_DIR_PATH, r"辅助程序/激光名称OCR修复规则.json")
    PRODUCT_ID_CATERGORY_CONVENTION_PATH = Path(PARENT_DIR_PATH, r"辅助程序/型号类别对照规则.json")
    GUI_GEOMETRY_PATH = Path(PARENT_DIR_PATH, r"辅助程序/程序几何.json")
    WORKPIECE_DICT = Path(PARENT_DIR_PATH, r"辅助程序/workpieceDict.json")
