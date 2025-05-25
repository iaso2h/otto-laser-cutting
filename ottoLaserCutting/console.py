logFlow = ""

def print(*args, **kwargs):
    import dearpygui.dearpygui as dpg
    global logFlow
    if logFlow == "":
        logFlow = "\n".join(args) + "\n"
    else:
        logFlow = logFlow + "\n".join(args) + "\n"
    dpg.set_value("log", value=logFlow)
