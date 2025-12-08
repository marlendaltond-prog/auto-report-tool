#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化报表工具整合脚本
用于检查、管理和使用AutoReport Pro的所有功能
"""

import os
import sys
import subprocess
import shutil
import json
import yaml
from datetime import datetime

# 项目根目录
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# 核心文件列表
CORE_FILES = [
    "auto_report.py",
    "report_gui.py",
    "run_gui.py",
    "requirements.txt",
    "config.yaml",
    "自动化报表工具使用说明书(优化版).md",
    "自动化报表工具使用说明书.docx"
]

# 辅助脚本列表
AUX_SCRIPTS = [
    "convert_md_to_docx.py",
    "create_simple_docx.py",
    "create_usage_docx.py",
    "create_usage_word.ps1"
]

# 依赖包列表
REQUIRED_PACKAGES = [
    "pandas",
    "numpy",
    "openpyxl",
    "sqlalchemy",
    "jinja2",
    "reportlab",
    "requests",
    "pillow",
    "pyyaml"
]


def check_file_exists(file_path):
    """检查文件是否存在"""
    return os.path.exists(file_path)


def check_all_files():
    """检查所有核心文件是否存在"""
    print("=== 检查核心文件 ===")
    missing_files = []
    for file in CORE_FILES:
        file_path = os.path.join(ROOT_DIR, file)
        if check_file_exists(file_path):
            print(f"✓ {file} - 存在")
        else:
            print(f"✗ {file} - 缺失")
            missing_files.append(file)
    
    print("\n=== 检查辅助脚本 ===")
    for file in AUX_SCRIPTS:
        file_path = os.path.join(ROOT_DIR, file)
        if check_file_exists(file_path):
            print(f"✓ {file} - 存在")
        else:
            print(f"✗ {file} - 缺失")
            missing_files.append(file)
    
    return missing_files


def check_dependencies():
    """检查依赖包是否安装"""
    print("\n=== 检查依赖包 ===")
    missing_packages = []
    
    # 使用pip list命令检查已安装的包
    try:
        import subprocess
        result = subprocess.run([sys.executable, "-m", "pip", "list", "--format=freeze"], 
                               capture_output=True, text=True, check=True)
        installed_packages = {line.split("==")[0].lower() for line in result.stdout.strip().splitlines()}
    except Exception as e:
        print(f"检查依赖包时出错: {e}")
        # 回退到import方法
        for package in REQUIRED_PACKAGES:
            try:
                __import__(package)
                print(f"✓ {package} - 已安装")
            except ImportError:
                print(f"✗ {package} - 未安装")
                missing_packages.append(package)
        return missing_packages
    
    # 检查每个需要的包
    for package in REQUIRED_PACKAGES:
        if package.lower() in installed_packages:
            print(f"✓ {package} - 已安装")
        else:
            print(f"✗ {package} - 未安装")
            missing_packages.append(package)
    
    return missing_packages


def install_dependencies():
    """安装所有依赖包"""
    print("\n=== 安装依赖包 ===")
    requirements_path = os.path.join(ROOT_DIR, "requirements.txt")
    if check_file_exists(requirements_path):
        cmd = [sys.executable, "-m", "pip", "install", "-r", requirements_path]
        print(f"执行命令: {' '.join(cmd)}")
        subprocess.run(cmd, check=True)
        print("依赖包安装完成")
    else:
        # 如果requirements.txt不存在，直接安装所需包
        for package in REQUIRED_PACKAGES:
            cmd = [sys.executable, "-m", "pip", "install", package]
            print(f"执行命令: {' '.join(cmd)}")
            subprocess.run(cmd, check=True)
        print("依赖包安装完成")


def run_gui():
    """运行GUI界面"""
    print("\n=== 启动GUI界面 ===")
    gui_path = os.path.join(ROOT_DIR, "run_gui.py")
    if check_file_exists(gui_path):
        subprocess.run([sys.executable, gui_path])
    else:
        print("✗ run_gui.py - 缺失，无法启动GUI")


def run_cli_example():
    """运行命令行示例"""
    print("\n=== 运行命令行示例 ===")
    main_script = os.path.join(ROOT_DIR, "auto_report.py")
    example_data = os.path.join(ROOT_DIR, "302594156_按序号_大学生对新能源汽车购买意向调查研究_254_246.xlsx")
    
    if check_file_exists(main_script) and check_file_exists(example_data):
        cmd = [
            sys.executable, main_script,
            "--data", example_data,
            "--output", "reports",
            "--format", "excel,pdf"
        ]
        print(f"执行命令: {' '.join(cmd)}")
        subprocess.run(cmd)
    else:
        print("✗ 示例数据或主脚本缺失")


def generate_config():
    """生成示例配置文件"""
    print("\n=== 生成配置文件 ===")
    config_path = os.path.join(ROOT_DIR, "report_config.json")
    
    config = {
        "report": {
            "name": "自动化报表",
            "title": "数据分析报告"
        },
        "data": {
            "type": "excel",
            "path": "302594156_按序号_大学生对新能源汽车购买意向调查研究_254_246.xlsx"
        },
        "output": {
            "dir": "reports",
            "formats": ["excel", "pdf"]
        }
    }
    
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    
    print(f"✓ 配置文件已生成: {config_path}")


def convert_md_to_docx():
    """将Markdown文档转换为Word文档"""
    print("\n=== 转换文档格式 ===")
    md_path = os.path.join(ROOT_DIR, "自动化报表工具使用说明书(优化版).md")
    docx_path = os.path.join(ROOT_DIR, "自动化报表工具使用说明书.docx")
    
    if check_file_exists(md_path):
        converter_path = os.path.join(ROOT_DIR, "convert_md_to_docx.py")
        if check_file_exists(converter_path):
            cmd = [sys.executable, converter_path, md_path, docx_path]
            print(f"执行命令: {' '.join(cmd)}")
            subprocess.run(cmd)
        else:
            print("✗ 转换脚本缺失")
    else:
        print("✗ Markdown文档缺失")


def create_output_dir():
    """创建输出目录"""
    output_dir = os.path.join(ROOT_DIR, "reports")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"\n=== 输出目录已创建: {output_dir} ===")


def main():
    """主函数"""
    print("=" * 50)
    print("自动化报表工具 - AutoReport Pro 整合脚本")
    print("=" * 50)
    print(f"当前目录: {ROOT_DIR}")
    print(f"执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # 1. 检查所有文件
    missing_files = check_all_files()
    
    # 2. 检查依赖
    missing_packages = check_dependencies()
    
    # 3. 安装缺失的依赖
    if missing_packages:
        print(f"\n发现 {len(missing_packages)} 个缺失的依赖包")
        print("自动安装缺失的依赖包...")
        install_dependencies()
    
    # 4. 创建输出目录
    create_output_dir()
    
    # 5. 主菜单
    print("\n" + "=" * 50)
    print("主菜单")
    print("=" * 50)
    print("1. 运行命令行示例")
    print("2. 启动GUI界面")
    print("3. 生成配置文件")
    print("4. 转换文档格式")
    print("5. 退出")
    
    choice = input("请选择操作 (1-5): ")
    
    if choice == '1':
        run_cli_example()
    elif choice == '2':
        run_gui()
    elif choice == '3':
        generate_config()
    elif choice == '4':
        convert_md_to_docx()
    elif choice == '5':
        print("退出整合脚本")
    else:
        print("无效选择")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程序被中断")
    except Exception as e:
        print(f"\n\n发生错误: {e}")
    finally:
        print("\n整合脚本执行完毕")
        input("按回车键退出...")