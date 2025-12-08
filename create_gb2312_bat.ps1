# 创建GB2312编码的中文启动批处理文件
$content = @'
@echo off
setlocal enabledelayedexpansion

echo 启动报表工具...
echo.

REM 检查Python是否安装
echo 检查Python安装...
python --version 2>NUL
if %ERRORLEVEL% NEQ 0 (
    echo 错误：Python未安装或未添加到PATH
    echo.
    echo 请按照以下步骤安装Python：
    echo 1. 从https://www.python.org/下载Python 3.7或更高版本
    echo 2. 运行安装程序并勾选"Add Python to PATH"
    echo 3. 点击"Install Now"完成安装
    echo.
    echo 详细指南请查看"Python安装指南.txt"
    echo.
    echo 按任意键退出...
    pause >NUL
    exit /B 1
)

REM 检查pip是否可用
echo 检查pip...
pip --version 2>NUL
if %ERRORLEVEL% NEQ 0 (
    echo 错误：pip不可用
    echo 请安装pip或更新Python
    echo.
    echo 按任意键退出...
    pause >NUL
    exit /B 1
)

REM 安装依赖
echo 安装依赖...
pip install pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow

REM 启动应用
echo 启动应用...
python auto_report.py

echo 应用已关闭
pause >NUL
endlocal
'@

# 使用UTF8-NOBOM编码保存，让cmd能正确识别中文
$content | Out-File -FilePath ".\启动报表工具.bat" -Encoding ASCII -Force

Write-Host "中文启动批处理文件已创建"