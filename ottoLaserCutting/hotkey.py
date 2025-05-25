import keySet
import cutRecord

from pynput import keyboard
from pynput.mouse import Button, Controller
from PIL import ImageGrab
import win32gui, win32process, win32api
import psutil
import time
import copy

mouse = Controller()
MOUSE_INTERVAL = 0.075


def hotkeyAlignTube():
    hwnd = win32gui.GetForegroundWindow()
    pId = win32process.GetWindowThreadProcessId(hwnd)[1]
    pName = psutil.Process(pId).name()
    if pName == "TubePro.exe":
        rect = win32gui.GetWindowRect(hwnd)
        winX = rect[0]
        winY = rect[1]
        img = ImageGrab.grab()
        imgRGB = img.convert("RGB")
        startPixelPos1    = [827, 1013]
        startPixelPos2    = [853, 1013]
        pausePixelPos    = [972, 1013]
        continuePixelPos = [850, 1077]
        stopPixelPos     = [972, 1077]
        if win32api.GetSystemMetrics(0) < win32api.GetSystemMetrics(1):
            targetCompletedPixel = imgRGB.getpixel((15, 1810))
        else:
            targetCompletedPixel = (255, 155, 155)

        if targetCompletedPixel == (170, 170, 0) or targetCompletedPixel == (255, 155, 155):
            deltaY = 60
        else:
            deltaY = 0

        startPixelPos1[1]   = startPixelPos1[1] + deltaY
        startPixelPos2[1]   = startPixelPos2[1] + deltaY
        pausePixelPos[1]    = pausePixelPos[1] + deltaY
        continuePixelPos[1] = continuePixelPos[1] + deltaY
        stopPixelPos[1]     = stopPixelPos[1] + deltaY
        startPixel1 = imgRGB.getpixel(startPixelPos1)
        startPixel2 = imgRGB.getpixel(startPixelPos2)
        pausePixel    = imgRGB.getpixel(pausePixelPos)
        continuePixel = imgRGB.getpixel(continuePixelPos)
        stopPixel     = imgRGB.getpixel(stopPixelPos)

        # if pausePixel == (255, 167, 51) and \
        #     continuePixel == (196, 196, 196) and \
        #     stopPixel == (228, 28, 28) and \
        #     ((startPixel1 == (0, 160, 45) and startPixel2 == (192, 192, 192)) or (startPixel1 == (192, 192, 192) and startPixel2 == (0, 160, 45))):
        savedPosition = copy.copy(mouse.position)
        mouse.position = (302, 96 + deltaY)
        mouse.press(Button.left)
        mouse.release(Button.left)
        time.sleep(MOUSE_INTERVAL)
        mouse.position = (302, 146 + deltaY)
        mouse.press(Button.left)
        mouse.release(Button.left)
        time.sleep(MOUSE_INTERVAL)
        mouse.position = (455, 173 + deltaY)
        mouse.press(Button.left)
        mouse.release(Button.left)
        mouse.position = savedPosition
        # else:
        #     print("-----------------------")
        #     print("激光机运行时无法进行此操作")
        #     print([ 0, 160, 45 ])
        #     print([ 192, 192, 192 ])
        #     print(f"pausePixel == (255, 167, 51): {pausePixel == (255, 167, 51)}")
        #     print(f"continuePixel == (196, 196, 196) = {continuePixel == (196, 196, 196)}")
        #     print(f"stopPixel == (228, 28, 28) = {stopPixel == (228, 28, 28)}")
        #     print(f"startPixel1 = {startPixel1}")
        #     print(f"startPixel2 = {startPixel2}")
        #     print(f"pausePixel = {pausePixel}")
        #     print(f"continuePixel = {continuePixel}")
        #     print(f"stopPixel = {stopPixel}")
            # return win32api.MessageBox(
            #         None,
            #         "激光机运行时无法进行此操作",
            #         "Warning",
            #         4096 + 0 + 16
            #     )
            # MB_SYSTEMMODAL==4096
            # Button Styles:
            # 0:OK  --  1:OK|Cancel -- 2:Abort|Retry|Ignore -- 3:Yes|No|Cancel -- 4:Yes|No -- 5:Retry|No -- 6:Cancel|Try Again|Continue
            # To also change icon, add these values to previous number
            # 16 Stop-sign  ### 32 Question-mark  ### 48 Exclamation-point  ### 64 Information-sign ('i' in a circle)


def coordinateEcho():
    return win32api.MessageBox(
            None,
            f"x: {mouse.position[0]}, y: {mouse.position[1]}",
            "Info",
            4096 + 0 + 64
        )
    # MB_SYSTEMMODAL==4096
    # Button Styles:
    # 0:OK  --  1:OK|Cancel -- 2:Abort|Retry|Ignore -- 3:Yes|No|Cancel -- 4:Yes|No -- 5:Retry|No -- 6:Cancel|Try Again|Continue
    # To also change icon, add these values to previous number
    # 16 Stop-sign  ### 32 Question-mark  ### 48 Exclamation-point  ### 64 Information-sign ('i' in a circle)


def onPress(key):
    if hasattr(key, "char"):
        keyName = key.char
    else:
        keyName = key.name
    if not keyName:
        return

    # Convert modifiers
    if "alt" in keyName:
        keyName = "alt"
    elif "ctrl" in keyName:
        keyName = "ctrl"
    elif "shift" in keyName:
        keyName = "shift"

    keySet.keys.add(keyName)

    if {"alt", "a"}.issubset(keySet.keys):
        hotkeyAlignTube()
    elif {"alt", "j"}.issubset(keySet.keys):
        cutRecord.takeScreenshot()


def onRelease(key):
    if hasattr(key, "char"):
        keyName = key.char
    else:
        keyName = key.name
    if not keyName:
        return

    # Convert modifiers
    if "alt" in keyName:
        keyName = "alt"
    elif "ctrl" in keyName:
        keyName = "ctrl"
    elif "shift" in keyName:
        keyName = "shift"

    if keyName in keySet.keys:
        keySet.keys.remove(keyName)

        # Force reset keySet if necessary
        if "ctrl" == keyName or "shift" == keyName or "alt" == keyName:
            if "ctrl" not in keySet.keys and "shift" not in keySet.keys and "alt" not in keySet.keys:
                keySet.keys = set()
