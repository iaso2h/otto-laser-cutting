import config
import rtfParse
import hotkey
import gui

import argparse


if __name__ == "__main__":
    argParser = argparse.ArgumentParser()
    argParser.add_argument("-D", "--dev", action="store_true")
    argParser.add_argument("-R", "--rtf", action="store_true")
    args = argParser.parse_args()
    if args.rtf:
        rtfParse.parseAllLog()
    else:
        listener = hotkey.keyboard.Listener(
            on_press=hotkey.onPress, on_release=hotkey.onRelease
        )
        listener.start()
        gui.dpg.show_viewport()
        gui.dpg.start_dearpygui()
        gui.dpg.destroy_context()
