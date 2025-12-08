@echo off
setlocal enabledelayedexpansion

echo Recreating virtual environment...

REM Remove existing virtual environment
if exist ".venv" (
    echo Removing existing virtual environment...
    rmdir /s /q ".venv"
)

REM Create new virtual environment
echo Creating new virtual environment...
python -m venv .venv

REM Activate virtual environment and install dependencies
echo Activating virtual environment and installing dependencies...
.venv\Scripts\activate.bat && pip install pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow cryptography

echo Virtual environment recreated successfully.
echo Testing installation...
.venv\Scripts\python.exe -c "from cryptography.fernet import Fernet; print('Cryptography imported successfully!')"

echo.
echo Press any key to exit...
pause >NUL