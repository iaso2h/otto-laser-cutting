# File: parseTubeProLog
# Author: iaso2h
# Description: Parsing Log files(.rtf) from TubePro and split them into separated files
VERSION     = "0.0.147"
LASTUPDATED = "06/06/2025"

import sys
import locale
import json
import re
from pathlib import Path
from dataclasses import dataclass, InitVar, field
from typing import Optional, cast
from datetime import datetime

@dataclass
class Geometry:
    xPos: int
    yPos: int
    width: int
    height: int

@dataclass
class Paths:
    otto:         Path = field(init=False)
    warehousing:  Path = field(init=False)
    _otto:        InitVar[str]
    _warehousing: InitVar[str]

    def __post_init__(self, _otto, _warehousing):
        self.otto        = Path(_otto)
        if not self.otto.exists():
            print(f"{self.otto} doesn't exist.")
            sys.exit()
        self.warehousing = Path(_warehousing)
        if not self.warehousing.exists():
            print(f"{self.warehousing} doesn't exist.")

@dataclass
class Pats:
    laserFile:           re.Pattern = field(init=False)
    workpieceDimension:  re.Pattern = field(init=False)
    _laserFile:          InitVar[str]
    _workpieceDimension: InitVar[str]

    def __post_init__(self, _laserFile: str, _workpieceDimension: str):
        # Compile regex patterns during initialization
        self.laserFile = re.compile(_laserFile)
        self.workpieceDimension = re.compile(_workpieceDimension)

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
# Load JSON and convert to dataclass
with open(EXTERNAL_CONFIG, "r", encoding="utf-8") as f:
    data = json.load(f)
    cfg = Configuration(
        geometry=Geometry(**data["geometry"]),
        fontSize=data["fontSize"],
        paths=Paths(**{
            "_otto":        data["paths"]["otto"],
            "_warehousing": data["paths"]["warehousing"],
        }),
        # paths=Paths(**data["paths"]),
        patterns=Pats(**{
            "_laserFile":          data["patterns"]["laserFile"],
            "_workpieceDimension": data["patterns"]["workpieceDimension"],
        }),
        # patterns=Pats(**data["patterns"]),
        email=Email(**data["email"])
    )

LASER_FILE_DIR_PATH  = Path(cfg.paths.otto, r"切割文件")
