from ottoLaserCutting.config import cfg

import shutil
import PyInstaller.__main__
from pathlib import Path
EXECUTABLE_PATH = Path(cfg.paths.otto, r"辅助程序/ottoLaserCutting")
PyInstaller.__main__.run(
    [
        "ottoLaserCutting.spec",
        "--distpath=" + str(EXECUTABLE_PATH),
        "--noconfirm",
        "--clean",
    ]
)

shutil.copy2(
    EXECUTABLE_PATH,
    Path(
        EXECUTABLE_PATH.parent,
        EXECUTABLE_PATH.stem + "Template" + EXECUTABLE_PATH.suffix
    )
)
