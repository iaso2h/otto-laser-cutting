import subprocess
import util
import config
from config import cfg
import cutRecord
import workpiece
import rtfParse

import os
import dearpygui.dearpygui as dpg
import win32api, win32con
from datetime import datetime, timedelta
from typing import Optional

dpg.create_context()
reg = dpg.add_font_registry()
fontName = dpg.add_font(file=r"C:\Windows\Fonts\msyh.ttc", size=cfg.fontSize, parent=reg)
dpg.add_font_range(0x0001, 0x9FFF, parent=fontName)
dpg.bind_font(fontName)

dpg.create_viewport(
        title="OttoLaserCutting",
        decorated=False,
        x_pos=cfg.geometry.xPos,
        y_pos=cfg.geometry.yPos,
        width=cfg.geometry.width,
        height=cfg.geometry.height,
        always_on_top=False,
        resizable=False,
    )

dpg.setup_dearpygui()


# Initiate GUI components
with dpg.window(
        label="欧拓开料辅助 v" + config.VERSION,
        autosize=False,
        no_resize=True,
        width=cfg.geometry.width,
        no_close=True,
        no_title_bar=False,
        no_move=True,
        no_collapse=True,
    ):
    loginName = os.getlogin()
    with dpg.group(horizontal=True, horizontal_spacing=60):
        dpg.add_text(f"编程: 阮焕")
        with dpg.tooltip(dpg.last_item()):
            dpg.add_text(f"OS User Name: {loginName}")
        dpg.add_text(f"最后更新: {config.LASTUPDATED}")
    dpg.add_separator(label="开料")
    with dpg.group(horizontal=True):
        # dpg.add_button(label="程序截图", callback=lambda _: cutRecord.takeScreenshot())
        with dpg.tooltip(dpg.last_item()):
            dpg.add_text("Ctrl-左键: 打开截图目录\nShift-左键: 重新链接截图")
        dpg.add_button(label="耗时分析", callback=rtfParse.parsePeriodLog)
        with dpg.tooltip(dpg.last_item()):
            dpg.add_text("Ctrl-左键: 打开日志目录\nShift-左键: 35天内耗时分析\nAlt-左键: 365天内耗时分析\nCtrl+Shift+Alt-左键: 历史切割耗时分析(单文件)")
        dpg.add_button(label="日志分析", callback=rtfParse.rtfSimplify)
        with dpg.tooltip(dpg.last_item()):
            dpg.add_text("Ctrl-左键: 打开日志目录\nShift-左键: 35天内日志分析\nAlt-左键: 365天内耗时分析")
    dpg.add_separator(label="排样文件")
    with dpg.group(horizontal=True):
        dpg.add_button(label="命名检查",     callback=workpiece.workpieceNamingVerification)
        with dpg.tooltip(dpg.last_item()):
            dpg.add_text("Ctrl-左键: 打开切割文件目录/nShift-左键: 仅检查.zx/.zzx激光图纸文件")
        dpg.add_button(label="工件规格总览", callback=workpiece.exportDimensions)
        with dpg.tooltip(dpg.last_item()):
            dpg.add_text("Ctrl-左键: 打开工件规格总览目录\nShift-左键: 仅导出.zx/.zzx激光图纸文件")
        dpg.add_button(label="删除冗余排样", callback=workpiece.removeRedundantLaserFile)
        with dpg.tooltip(dpg.last_item()):
            dpg.add_text("Ctrl-左键: 打开切割文件目录")
    dpg.add_separator(label="开料实时检测")
    with dpg.group(horizontal=True):
        laserMonitorToggle        = dpg.add_button(label="切换检测")
        laserMonitorCheckTemplate = dpg.add_button(label="匹配检测")
    dpg.add_input_text(
        multiline=True,
        default_value="",
        tab_input=True,
        tracked=False,
        track_offset=0,
        width=cfg.geometry.width - 30,
        height=155,
        readonly=True,
        tag="log",
        no_horizontal_scroll=False,
    )
    def clearLog():
        util.logFlow.clear()
        dpg.set_value("log", value=util.logFlow)
    def toggleLogTrackihg(sender, appData):
        currentTracked = dpg.get_item_configuration("log")["tracked"]
        dpg.configure_item("log", tracked=(not currentTracked))

    with dpg.group(horizontal=True):
        # BUG:
        # dpg.add_checkbox(label="自动滚动日志", default_value=False, callback=toggleLogTrackihg)
        dpg.add_button(label="清除日志", callback=clearLog)

    with dpg.group(horizontal=True):
        dpg.add_button(label="退出", callback=dpg.destroy_context)
        shutdownNotification = dpg.add_text(label="placeHolder")
        shutdownPicker = dpg.add_time_picker(
                label="timePicker",
                hour24=True,
                default_value={
                    "hour": config.LAUNCH_TIME.hour,
                    "min": config.LAUNCH_TIME.minute,
                    "sec": config.LAUNCH_TIME.second,
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



import tubeProMonitor
tubeProMonitor.monitor = tubeProMonitor.Monitor()
dpg.configure_item(
        laserMonitorToggle,
        callback=tubeProMonitor.monitor.toggleMonitoring
        )
dpg.configure_item(
        laserMonitorCheckTemplate,
        callback=tubeProMonitor.monitor.checkTemplateMatches
        )
