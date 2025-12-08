#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
直接启动自动化报表工具的Python脚本
"""

import os
import sys
import subprocess
import traceback

def check_python_version():
    """检查Python版本"""
    print(f"Python版本: {sys.version}")
    print(f"Python路径: {sys.executable}")
    if sys.version_info < (3, 7):
        print("警告: Python版本过低，建议使用Python 3.7或更高版本")
        return False
    return True

def check_required_files():
    """检查必要文件是否存在"""
    required_files = ["run_gui.py", "report_gui.py"]
    all_exists = True
    
    print("\n检查必要文件:")
    for file in required_files:
        if os.path.exists(file):
            print(f"✓ {file} 存在")
        else:
            print(f"✗ {file} 不存在")
            all_exists = False
    
    return all_exists

def check_dependencies():
    """检查必要的依赖是否安装"""
    required_deps = [
        "pandas", "openpyxl", "sqlalchemy", "jinja2", 
        "reportlab", "requests", "pyyaml", "pillow"
    ]
    
    print("\n检查依赖包:")
    missing_deps = []
    
    for dep in required_deps:
        try:
            __import__(dep)
            print(f"✓ {dep} 已安装")
        except ImportError:
            print(f"✗ {dep} 未安装")
            missing_deps.append(dep)
    
    if missing_deps:
        print(f"\n缺失的依赖: {', '.join(missing_deps)}")
        return False
    
    return True

def install_dependencies():
    """安装必要的依赖"""
    required_deps = [
        "pandas", "openpyxl", "sqlalchemy", "jinja2", 
        "reportlab", "requests", "pyyaml", "pillow"
    ]
    
    print("\n安装依赖包:")
    
    try:
        import pip
        pip.main(["install", "-U"] + required_deps)
        print("依赖安装完成")
        return True
    except Exception as e:
        print(f"依赖安装失败: {e}")
        return False

def start_application():
    """启动自动化报表工具"""
    print("\n启动自动化报表工具...")
    
    try:
        from report_gui import ReportGUI
        import tkinter as tk
        
        root = tk.Tk()
        app = ReportGUI(root)
        root.mainloop()
        
        print("应用程序已正常退出")
        return True
    except Exception as e:
        print(f"应用程序启动失败: {e}")
        print("详细错误信息:")
        traceback.print_exc()
        return False

def main():
    """主函数"""
    print("="*50)
    print("自动化报表工具启动器")
    print("="*50)
    
    # 1. 检查Python版本
    if not check_python_version():
        input("\n按回车键退出...")
        return
    
    # 2. 检查必要文件
    if not check_required_files():
        input("\n按回车键退出...")
        return
    
    # 3. 检查依赖
    if not check_dependencies():
        choice = input("\n是否安装缺失的依赖? (y/n): ")
        if choice.lower() == 'y':
            if not install_dependencies():
                input("\n按回车键退出...")
                return
        else:
            print("跳过依赖安装")
    
    # 4. 启动应用程序
    start_application()
    
    print("\n" + "="*50)
    input("按回车键退出...")

if __name__ == "__main__":
    main()
