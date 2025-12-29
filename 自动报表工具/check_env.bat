@echo off
chcp 936 >nul
echo 检查系统环境...
echo.
echo 1. 当前目录：
cd
set "CURRENT_DIR=%cd%"
echo %CURRENT_DIR%
echo.
echo 2. Python路径：
where python > python_path.txt 2>&1
if exist python_path.txt (
    type python_path.txt
) else (
    echo Python未找到
)
echo.
echo 3. 虚拟环境检查：
if exist ".venv\Scripts\python.exe" (
    echo 虚拟环境存在：%CURRENT_DIR%\.venv\Scripts\python.exe
    echo 虚拟环境版本：
    .venv\Scripts\python.exe --version 2>&1
) else (
    echo 虚拟环境不存在
)
echo.
echo 4. 文件检查：
if exist "run_gui.py" (
    echo run_gui.py 存在
) else (
    echo run_gui.py 不存在
)
if exist "report_gui.py" (
    echo report_gui.py 存在
) else (
    echo report_gui.py 不存在
)
echo.
echo 检查完成
pause