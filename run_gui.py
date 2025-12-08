#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化报表工具启动脚本
"""

import os
import sys
import subprocess
import traceback

def check_dependencies():
    """检查并安装必要的依赖"""
    required_deps = [
        "pandas", "openpyxl", "sqlalchemy", "jinja2", 
        "reportlab", "requests", "pyyaml", "pillow"
    ]
    
    print("检查依赖包...")
    missing = []
    
    for dep in required_deps:
        try:
            __import__(dep)
        except ImportError:
            missing.append(dep)
    
    if missing:
        print(f"安装缺失的依赖: {', '.join(missing)}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-U"] + missing)
            print("依赖安装完成")
            return True
        except subprocess.CalledProcessError:
            print("依赖安装失败")
            return False
    
    print("所有依赖已安装")
    return True

def main():
    """主函数"""
    print("自动化报表工具")
    print("="*30)
    
    # 检查当前目录
    print(f"当前目录: {os.getcwd()}")
    
    # 检查Python版本
    print(f"Python版本: {sys.version.split()[0]}")
    
    # 检查依赖
    if not check_dependencies():
        print("请手动安装依赖: pip install pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow")
        input("按回车键退出...")
        return
    
    # 导入并启动应用
    try:
        from report_gui import ReportGUI
        import tkinter as tk
        
        print("\n启动GUI界面...")
        root = tk.Tk()
        app = ReportGUI(root)
        root.mainloop()
        
    except ImportError as e:
        print(f"导入错误: {e}")
        print(traceback.format_exc())
    except Exception as e:
        print(f"启动错误: {e}")
        print(traceback.format_exc())
    
    print("应用已退出")

if __name__ == "__main__":
    main()