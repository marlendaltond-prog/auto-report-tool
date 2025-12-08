@echo off
echo Uninstall Report Tool...
echo.

REM Check if running in the correct directory
echo Check running directory...
if not exist "auto_report.py" (
    echo ERROR: Please run this uninstaller in the Report Tool folder!
    echo.
    pause
    exit /B 1
)

REM Confirm uninstallation
echo WARNING: This will uninstall the Report Tool
echo Type Y to continue, N to cancel:
set /p confirm="Confirm (Y/N): "
if /i "%confirm%" NEQ "Y" (
    echo Uninstallation canceled
    pause
    exit /B 0
)

echo.
echo Starting uninstallation...

REM Stop running Python processes
echo Stop running processes...
tasklist | findstr /i "python" >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    for /f "tokens=2" %%i in ('tasklist ^| findstr /i "python"') do (
        taskkill /f /pid %%i >NUL 2>&1
    )
)

REM Delete shortcuts
echo Delete desktop shortcuts...
del "%USERPROFILE%\Desktop\AutoReportTool.lnk" 2>NUL
del "%USERPROFILE%\Desktop\自动化报表工具.lnk" 2>NUL

echo Delete start menu shortcuts...
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AutoReportTool.lnk" 2>NUL

REM Clean temporary files
echo Clean temporary files...
del "*.log" 2>NUL
del "*.pyc" 2>NUL
del "*.whl" 2>NUL
rd /s /q "__pycache__" 2>NUL

REM Ask about removing dependencies
echo.
echo Remove Python dependencies?
set /p remove_deps="Remove dependencies (Y/N): "
if /i "%remove_deps%" EQU "Y" (
    echo Uninstall Python dependencies...
    pip --version >NUL 2>&1
    if %ERRORLEVEL% EQU 0 (
        pip uninstall -y pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow >NUL 2>&1
        echo Dependencies uninstalled
    ) else (
        echo WARNING: pip not available, skip dependency uninstall
    )
)

REM Ask about removing configuration files
echo.
echo Remove configuration files?
set /p remove_config="Remove config files (Y/N): "
if /i "%remove_config%" EQU "Y" (
    echo Delete configuration files...
    del "config.yaml" 2>NUL
    del "users.json" 2>NUL
    echo Configuration files deleted
)

REM Ask about removing virtual environment
if exist ".venv" (
    echo.
    echo Remove virtual environment?
    set /p remove_venv="Remove virtual env (Y/N): "
    if /i "%remove_venv%" EQU "Y" (
        echo Delete virtual environment...
        rd /s /q ".venv" 2>NUL
        echo Virtual environment deleted
    )
)

echo.
echo Uninstallation completed!
echo.
echo Removed:
echo - Desktop shortcuts
echo - Start menu shortcuts
echo - Temporary files
echo - Python dependencies (optional)
echo - Configuration files (optional)
echo - Virtual environment (optional)
echo.
echo To completely remove the tool, manually delete this folder:
echo %cd%
echo.
pause