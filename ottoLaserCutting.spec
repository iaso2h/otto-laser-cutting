from ottoLaserCutting import tubeProMonitor, config

import os
from pathlib import Path
from sys import version_info
from typing import Optional
envPathStr = os.environ.get('CONDA_PREFIX')
if not envPathStr:
    raise RuntimeError("Not in Conda environment")
envPath = Path(envPathStr)
condaPath = envPath.parent.parent
pythonVersion = str(version_info.major) + "." + str(version_info.minor)


# Find MKL dll paths
condaPkgPath = Path(condaPath, "pkgs")
resultPaths = list(condaPkgPath.glob("mkl*/Library/bin"))
if not resultPaths:
    raise RuntimeError("Cannot find MKL dll", resultPaths)

# Use the newest ones
mklNewest:Optional[Path] = None
for p in resultPaths:
    if not mklNewest:
        mklNewest = p
    else:
        if p.stat().st_mtime > mklNewest.stat().st_mtime:
            mklNewest = p

if not mklNewest:
    raise RuntimeError("Cannot find MKL dll")


binariesArg = []

# Hanldle inter Pyinstaller numpy "Intel MKL FATAL ERROR: Cannot load mkl_intel_thread.dll"
# Reference: https://stackoverflow.com/questions/35478526/pyinstaller-numpy-intel-mkl-fatal-error-cannot-load-mkl-intel-thread-dll
for p in mklNewest.iterdir():
    if p.stem.startswith("mkl"):
        binariesArg.append((str(p), "."))


# Handle Error “ImportError: ERROR: recursion is detected during loading of “cv2” binary extensions. Check OpenCV installation.” with Pyinstaller
# Credit: https://github.com/opencv/opencv-python/issues/680
cv2Path = Path(envPath, "Lib", "site-packages", "cv2")
binariesArg.append(
    (
        str(
            Path(cv2Path, f"python-{pythonVersion}")),
            f"./cv2/python-{pythonVersion}"
    )
)

# Add monitor matching templates
for pic in tubeProMonitor.PIC_TEMPLATE.iterdir():
    if pic.is_file() and pic.suffix == ".png":
        binariesArg.append(
            (
                str(pic),
                "./" + str(pic.parent.relative_to(config.EXECUTABLE_DIR))
            )
        )


block_cipher = None

a = Analysis( # type: ignore
    ['ottoLaserCutting/__main__.py'],
    pathex=[str(Path(cv2Path, f"python-{pythonVersion}"))],
    binaries=binariesArg,
    datas=[],
    hiddenimports=['openpyxl.cell._writer'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher) # type: ignore

exe = EXE( # type: ignore
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OttoLaserCutting',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='./src/otto.ico',
)


# vim:ts=4:sts=4:sw=4:ft=python:fdm=marker
