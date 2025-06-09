import config
import rtfParse
import hotkey
import gui
import tubeProMonitor

import os
import sys
import argparse


if __name__ == "__main__":
    argParser = argparse.ArgumentParser()
    argParser.add_argument("-D", "--dev", action="store_true")
    args = argParser.parse_args()
    listener = hotkey.keyboard.Listener(
        on_press=hotkey.onPress, on_release=hotkey.onRelease
    )
    listener.start()
    if os.getlogin() != "OT03":
        tubeProMonitor.monitor.toggleMonitoring() # type: ignore

    gui.dpg.show_viewport()
    gui.dpg.start_dearpygui()
    gui.dpg.destroy_context()
