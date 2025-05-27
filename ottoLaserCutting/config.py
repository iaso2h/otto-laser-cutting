# File: parseTubeProLog
# Author: iaso2h
# Description: Parsing Log files(.rtf) from TubePro and split them into separated files
VERSION     = "0.0.132"
LASTUPDATED = "2025-05-27"

import sys
import locale
import json
from pathlib import Path
from dataclasses import dataclass
from typing import Optional
from datetime import datetime
locale.setlocale(locale.LC_TIME, '')
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    BUNDLE_MODE = True
    BUNDLE_PATH = Path(sys._MEIPASS)  # type: ignore
    EXECUTABLE_DIR = Path(sys.executable).parent
else:
    BUNDLE_MODE = False
    EXECUTABLE_DIR = Path(__file__).parent.parent
EXTERNAL_CONFIG = Path(EXECUTABLE_DIR, "configuration.json")
if not EXTERNAL_CONFIG.exists():
    raise FileExistsError(f"Can't find external configuration at: {str(EXTERNAL_CONFIG)}.")
LAUNCH_TIME = datetime.now()

@dataclass
class Geometry:
    xPos: int
    yPos: int
    width: int
    height: int

@dataclass
class Paths:
    otto: str
    warehousing: str

@dataclass
class Pats:
    laserFile: str
    workpieceDimension: str

@dataclass
class Email:
    sslPort: int
    smtpServer: str
    senderAccount: str
    senderPassword: str
    receiverAccounts: list[str]

@dataclass
class Configuration:
    geometry: Geometry
    fontSize: int
    paths: Paths
    patterns: Pats
    email: Email


# Load JSON and convert to dataclass
with open(EXTERNAL_CONFIG, "r", encoding="utf-8") as f:
    data = json.load(f)
    cfg = Configuration(
        geometry=Geometry(**data["geometry"]),
        fontSize=data["fontSize"],
        paths=Paths(**data["paths"]),
        patterns=Pats(**data["patterns"]),
        email=Email(**data["email"])
    )


if not Path(cfg.paths.otto).exists():
    print('无法找到"欧拓图纸"文件夹')
    sys.exit()

LASER_FILE_DIR_PATH  = Path(cfg.paths.otto, r"切割文件")
