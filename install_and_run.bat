@echo off
echo Installing dependencies...
python -m pip install pandas openpyxl sqlalchemy jinja2 reportlab requests pyyaml pillow cryptography
echo.
echo Testing cryptography import...
python -c "from cryptography.fernet import Fernet; print('Cryptography imported successfully!')"
echo.
echo Starting application...
python auto_report.py
echo.
echo Application closed.
pause >NUL