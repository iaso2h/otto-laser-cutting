import chardet
import re
from striprtf.striprtf import rtf_to_text
from pathlib import Path

partLogParentDir = Path(r"D:\Program Files\Git\Repos\parse-tube-cutting-log\parseTubeCuttingLog")

def iterCopy(partLogParentDir: Path) -> list:
    rtfAll = []
    txtAll = []
    for srcPath in partLogParentDir.iterdir():
        if srcPath.is_dir() and re.search(r"^\d", srcPath.name):
            rtfSub, txtSub = iterCopy(srcPath)
            rtfAll.extend(rtfSub)
            txtAll.extend(txtSub)
        else:
            if srcPath.suffix == ".txt":
                txtAll.append(srcPath)
            elif srcPath.suffix == ".rtf":
                rtfAll.append(srcPath)

    return rtfAll, txtAll

rtfAll, txtAll = iterCopy(partLogParentDir)


def getEncoding(filePath):
    # Create a magic object
    with open(filePath, "rb") as f:
        # Detect the encoding
        rawData = f.read()
        result = chardet.detect(rawData)
        return result["encoding"]

for rtf in rtfAll:
    myFile = str(rtf)
    rtfEncoding = getEncoding(myFile)

    with open(myFile, "rb") as f:
        decodedContent = rtf_to_text(f.read().decode(rtfEncoding))
    with open(myFile, "w", encoding="utf-8") as f:
        f.write(decodedContent)
