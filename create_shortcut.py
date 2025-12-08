#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建自动化报表工具图形界面的快捷方式
"""

import os
import sys
from win32com.client import Dispatch

# 定义路径
startup_script = r"C:\Users\25331\Desktop\新建文件夹\启动报表工具.bat"
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
shortcut_path = os.path.join(desktop_path, "AutoReportTool.lnk")

# 检查启动脚本是否存在
if not os.path.exists(startup_script):
    print("错误：启动脚本不存在！")
    input("按任意键退出...")
    sys.exit(1)

# 创建快捷方式
try:
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = startup_script
    shortcut.WorkingDirectory = r"C:\Users\25331\Desktop\新建文件夹"
    shortcut.save()
    
    print(f"快捷方式已创建：{shortcut_path}")
    print("双击此快捷方式即可启动自动化报表工具图形界面！")
except Exception as e:
    print(f"创建快捷方式失败：{e}")

input("按任意键退出...")