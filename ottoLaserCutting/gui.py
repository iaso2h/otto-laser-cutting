import subprocess
import console
import config
import cutRecord
import workpiece
import rtfParse
import tubeProMonitor

import os
import json
import dearpygui.dearpygui as dpg
import win32api, win32con
from datetime import datetime, timedelta
from pprint import pprint

if win32api.GetSystemMetrics(0) < win32api.GetSystemMetrics(1) and config.GUI_GEOMETRY_PATH.exists():
    with open(config.GUI_GEOMETRY_PATH, "r", encoding="utf-8") as f:
        geo = json.load(f)
else:
    geo = {
            "x_pos": 800,
            "y_pos": 600,
            "width": 290,
            "height": 192,
            "fontSize": 16
    }
dpg.create_context()
reg = dpg.add_font_registry()
fontName = dpg.add_font(file=r"C:\Windows\Fonts\msyh.ttc", size=geo["fontSize"], parent=reg)
dpg.add_font_range(0x0001, 0x9FFF, parent=fontName)
dpg.bind_font(fontName)

dpg.create_viewport(
        title="ottoLaserCutting",
        decorated=False,
        x_pos=geo["x_pos"],
        y_pos=geo["y_pos"],
        width=geo["width"],
        height=geo["height"],
        always_on_top=False,
        resizable=False,
    )

dpg.setup_dearpygui()

with dpg.window(
        label="欧拓开料辅助 v" + config.VERSION,
        autosize=False,
        no_resize=True,
        width=geo["width"],
        no_close=True,
        no_title_bar=False,
        no_move=True,
        no_collapse=True,
    ):
    loginName = os.getlogin()
    with dpg.group(horizontal=True, horizontal_spacing=60):
        dpg.add_text(f"编程: 阮焕")
        with dpg.tooltip(dpg.last_item()):
            dpg.add_text(f"OS User Name: {loginName}\nDev Mode: {config.DEV_MODE}\nSilent Mode: {config.SILENT_MODE}")
        dpg.add_text(f"最后更新: {config.LASTUPDATED}")
    dpg.add_separator(label="开料")
    with dpg.group(horizontal=True):
        dpg.add_button(label="程序截图", callback=cutRecord.takeScreenshot)
        dpg.add_button(label="耗时分析", callback=rtfParse.parsePeriodLog)
        dpg.add_button(label="日志分析", callback=rtfParse.rtfSimplify)
    dpg.add_separator(label="排样文件")
    with dpg.group(horizontal=True):
        dpg.add_button(label="命名检查",     callback=workpiece.workpieceNamingVerification)
        dpg.add_button(label="工件规格总览", callback=workpiece.exportDimensions)
        dpg.add_button(label="删除冗余排样", callback=workpiece.removeRedundantLaserFile)
    dpg.add_separator(label="开料实时检测")
    with dpg.group(horizontal=True):
        tubeProMonitor.monitor = tubeProMonitor.Monitor()
        tubeProMonitor.monitor.loadTemplates()
        dpg.add_button(label="监视切割", callback=tubeProMonitor.monitor.toggleMonitoring)
        dpg.add_button(label="匹配检测", callback=tubeProMonitor.monitor.checkTemplateMatches)
    dpg.add_input_text(
        multiline=True,
        default_value=console.logFlow,
        tab_input=True,
        tracked=False,
        width=geo["width"] - 30,
        height=155,
        readonly=True,
        tag="log",
        no_horizontal_scroll=False,
    )
    def clearLog():
        console.logFlow = ""
        dpg.set_value("log", value=console.logFlow)

    with dpg.group(horizontal=True):
        dpg.add_button(label="退出", callback=dpg.destroy_context)
        guiStartTime = datetime.now()
        shutdownNotification = dpg.add_text(label="placeHolder")
        shutdownPicker = dpg.add_time_picker(
                label="timePicker",
                hour24=True,
                default_value={
                    "hour": guiStartTime.hour,
                    "min": guiStartTime.minute,
                    "sec": guiStartTime.second,
                    }
                )
        shutdownBtn = dpg.add_button(label="定时关机")
        def shutDownCallBack():
            shutDownVal = dpg.get_value(shutdownPicker)
            now = datetime.now()
            shutdownTime = now.replace(
                    hour=shutDownVal["hour"],
                    minute=shutDownVal["min"],
                    second=shutDownVal["sec"],
                    )
            if shutdownTime < now:
                shutdownTime = shutdownTime + timedelta(days=1)

            shutdownTimeReadable = datetime.strftime(shutdownTime, "%c")
            if win32con.IDYES == win32api.MessageBox(
                None,
                f"是否在{shutdownTimeReadable}关机？",
                "关机确认",
                4096 + 4 + 32
                ):
                #   MB_SYSTEMMODAL==4096
                ##  Button Styles:
                ### 0:OK  --  1:OK|Cancel -- 2:Abort|Retry|Ignore -- 3:Yes|No|Cancel -- 4:Yes|No -- 5:Retry|No -- 6:Cancel|Try Again|Continue
                ##  To also change icon, add these values to previous number
                ### 16 Stop-sign  ### 32 Question-mark  ### 48 Exclamation-point  ### 64 Information-sign ('i' in a circle)
                dpg.hide_item(shutdownPicker)
                dpg.hide_item(shutdownBtn)
                dpg.set_value(shutdownNotification, f"将于{shutdownTimeReadable}关机")
                secondsToShutdown = int((shutdownTime - now).total_seconds())
                subprocess.call(["shutdown", "-s", "-t", f"{secondsToShutdown}"])

        dpg.set_item_callback(shutdownBtn, shutDownCallBack)
        dpg.add_button(label="清除日志", callback=clearLog)


