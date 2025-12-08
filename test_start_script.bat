@echo off
setlocal enabledelayedexpansion

REM Redirect all output to a log file
echo %date% %time% - Starting test >> test_log.txt

REM Test virtual environment Python
if exist ".venv\Scripts\python.exe" (
    echo Found virtual environment Python >> test_log.txt
    ".venv\Scripts\python.exe" --version >> test_log.txt 2>&1
    ".venv\Scripts\pip.exe" --version >> test_log.txt 2>&1
    
    echo Installing dependencies... >> test_log.txt
    ".venv\Scripts\pip.exe" install pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow cryptography >> test_log.txt 2>&1
    
    echo Starting application... >> test_log.txt
    ".venv\Scripts\python.exe" auto_report.py >> test_log.txt 2>&1
    
    echo Application exited with code: !ERRORLEVEL! >> test_log.txt
) else (
    echo Virtual environment not found >> test_log.txt
)

echo %date% %time% - Test completed >> test_log.txt
echo Test completed. Check test_log.txt for details.
pause >NUL