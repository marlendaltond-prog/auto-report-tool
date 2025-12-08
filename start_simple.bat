@echo off
chcp 936 >nul
echo ==========================================
echo 自动化报表工具 - 简易启动器
echo ==========================================
echo.

REM 检查Python环境
echo 1. 检查Python环境...
python --version
if %ERRORLEVEL% NEQ 0 (
    echo 错误: 未找到Python解释器！
    echo 请确保已安装Python并将其添加到系统PATH中。
    echo.
    pause
    exit /B 1
)

echo ✓ Python环境正常。
echo.

REM 检查必要文件
echo 2. 检查必要文件...

if not exist "run_gui.py" (
    echo 错误: run_gui.py 文件不存在！
    pause
    exit /B 1
)

if not exist "report_gui.py" (
    echo 错误: report_gui.py 文件不存在！
    pause
    exit /B 1
)

echo ✓ 所有必要文件存在。
echo.

REM 检查并安装依赖
echo 3. 检查依赖包...
echo (正在检查pandas, openpyxl, sqlalchemy, jinja2, reportlab, requests, pyyaml, pillow)
echo.

REM 使用python -m pip确保使用正确的pip版本
python -m pip install -U pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo 警告: 依赖安装可能不完整！
    echo 请尝试手动运行以下命令:
    echo python -m pip install pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow
    echo.
    pause
)

echo.
echo ✓ 依赖检查完成。
echo.

REM 启动应用程序
echo 4. 启动自动化报表工具...
echo.
python run_gui.py

REM 检查应用程序退出状态
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo 错误: 应用程序启动失败！
    echo 请检查上面的错误信息。
    echo.
    pause
    exit /B 1
)

echo.
echo ✓ 应用程序已正常退出。
pause
