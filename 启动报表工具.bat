@echo off

REM Switch to script directory
cd /d "%~dp0"

REM Simple launcher script
setlocal enabledelayedexpansion

REM Display basic info
echo Automated Report Tool Launcher
echo =============================
echo.

REM Check required files
if not exist "run_gui.py" (
    echo ERROR: run_gui.py not found
    pause
    exit /b 1
)

if not exist "report_gui.py" (
    echo ERROR: report_gui.py not found
    pause
    exit /b 1
)

echo Starting Automated Report Tool...
echo.

REM Try virtual environment first
if exist ".venv\Scripts\python.exe" (
    echo Using virtual environment Python
    .venv\Scripts\python.exe run_gui.py
    goto end
)

REM If virtual environment doesn't exist, use system Python
echo Using system Python
python run_gui.py

:end
REM Check launch result
if %ERRORLEVEL% neq 0 (
    echo.
    echo Launch failed!
    echo Please make sure Python and all dependencies are installed.
    echo You can try running: python -m pip install pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow
)

echo.
echo Press any key to exit...
pause >nul