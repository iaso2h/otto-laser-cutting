import dearpygui.dearpygui as dpg
logFlow = ""


def print(*args, **kwargs):
    """
    Custom print function that redirects output to a Dear PyGui log window.
    Maintains a global log buffer (logFlow) and updates the GUI log display.

    Args:
        *args: Strings to be printed (will be joined with newlines).
        **kwargs: Unused, maintained for print() signature compatibility.
    """

    global logFlow
    if logFlow == "":
        logFlow = "\n".join(args) + "\n"
    else:
        logFlow = logFlow + "\n".join(args) + "\n"
    dpg.set_value("log", value=logFlow)
