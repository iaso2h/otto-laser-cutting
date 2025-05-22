import config

import PyInstaller.__main__
import os
from pathlib import Path
config.PROGRAM_DIR = Path(os.getcwd())

PyInstaller.__main__.run([
    "__main__.py",
    "--onefile",
    "--noconfirm",
    "--noconsole",
    # UGLY:
    "--add-binary=D:/miniconda3/pkgs/mkl-2024.2.2-h66d3029_15/Library/bin/mkl_intel_thread.2.dll:.",
    "--clean",
    "--distpath=" + str(Path(config.PARENT_DIR_PATH, "辅助程序")),
    "--name=TubeProAid",
    "--hidden-import=openpyxl.cell._writer",
    "--icon=./src/sticky-note.ico",
])
