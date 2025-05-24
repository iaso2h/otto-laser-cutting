import console
import config
config.updaPath()
import rtfParse
import hotkey
import gui

import argparse
import sys

if not config.PARENT_DIR_PATH.exists():
    import os
    cwd = os.getcwd()
    idx = cwd.find("欧拓图纸")
    if idx > -1:
        from pathlib import Path
        config.PARENT_DIR_PATH = Path(cwd[:idx+5])
        config.updaPath()
    else:
        import sys
        print('无法找到"欧拓图纸"文件夹')
        sys.exit()



print = console.print


if __name__ == "__main__":
    argParser = argparse.ArgumentParser()
    argParser.add_argument("-D", "--dev",    action="store_true")
    argParser.add_argument("-R", "--rtf",    action="store_true")
    args = argParser.parse_args()
    config.DEV_MODE = args.dev

    if args.rtf:
        rtfParse.parseAllLog()
    else:
        listener= hotkey.keyboard.Listener(
                on_press   = hotkey.onPress,
                on_release = hotkey.onRelease
            )
        listener.start()
        gui.dpg.show_viewport()
        gui.dpg.start_dearpygui()
        gui.dpg.destroy_context()
