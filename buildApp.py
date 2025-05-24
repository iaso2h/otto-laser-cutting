from ottoLaserCutting import config

import PyInstaller.__main__
import os
from pathlib import Path
# for p in Path(r"D:\miniconda3\pkgs\mkl-2024.2.2-h66d3029_15\Library\bin").iterdir():
#     if p.stem.startswith("mkl"):
#         args.append(f"--add-binary={str(p)}:.")

PyInstaller.__main__.run(
    [
        "ottoLaserCutting.spec",
        "--distpath=" + str(Path(config.cfg.paths.otto, r"辅助程序/ottoLaserCutting")),
        "--noconfirm",
        "--clean",
    ]
)
