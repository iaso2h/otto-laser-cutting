import config
logFlow = ""
richTextPat = r"\[[_/ a-zA-z]+?\]"

from rich.console import Console


console = Console()
def print(*args, **kwargs):
    if not config.GUI_MODE:
        if not config.SILENT_MODE:
            console.print(*args, **kwargs)
    else:
        import dearpygui.dearpygui as dpg
        import re
        global logFlow
        if logFlow == "":
            logFlow = "\n".join(args) + "\n"
        else:
            logFlow = logFlow + "\n".join(args) + "\n"
        logFlow = re.sub(richTextPat, "", logFlow)
        dpg.set_value("log", value=logFlow)

