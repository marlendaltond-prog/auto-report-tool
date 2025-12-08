@echo off
echo Checking Python availability...
echo.
python --version 2>&1
if errorlevel 1 (
    echo ERROR: Python is not found in PATH
    echo.
    echo Please install Python 3.7 or higher from https://www.python.org/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
)

python3 --version 2>&1
if errorlevel 1 (
    echo ERROR: Python3 is not found in PATH
    echo.
)

echo Checking common Python installation paths...
echo.
if exist "C:\Program Files\Python*" (
    dir "C:\Program Files\Python*" /b
) else (
    echo No Python found in C:\Program Files
)

if exist "C:\Program Files (x86)\Python*" (
    dir "C:\Program Files (x86)\Python*" /b
) else (
    echo No Python found in C:\Program Files (x86)
)

echo.
echo Press any key to exit...
pause >nul