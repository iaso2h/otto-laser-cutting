import sys
from pathlib import Path
sys.path.append(str(
    Path(
        Path(__file__).parent,
        "ottoLaserCutting"
        )
    )
)

from ottoLaserCutting.config import cfg, EXTERNAL_CONFIG
from ottoLaserCutting import tubeProMonitor

import shutil
import PyInstaller.__main__

EXPORT_EXECUTABLE_PATH = Path(cfg.paths.otto, r"辅助程序/OttoLaserCutting")
PyInstaller.__main__.run(
    [
        "ottoLaserCutting.spec",
        "--distpath=" + str(EXPORT_EXECUTABLE_PATH),
        "--noconfirm",
        "--clean",
    ]
)

# Ship external resource along with the bundle executable
shutil.copy2(
    EXTERNAL_CONFIG,
    Path(
        EXPORT_EXECUTABLE_PATH,
        EXTERNAL_CONFIG.stem + "Template" + EXTERNAL_CONFIG.suffix
    )
)

