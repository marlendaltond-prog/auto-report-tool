@echo off
echo Testing system environment...
echo.
echo Current directory: %cd%
echo.
echo Python check:
python --version 2>NUL
if %ERRORLEVEL% NEQ 0 (
    echo Python not found in PATH
) else (
    echo Python found
)
echo.
echo Pip check:
pip --version 2>NUL
if %ERRORLEVEL% NEQ 0 (
    echo Pip not found in PATH
) else (
    echo Pip found
)
echo.
echo PATH environment variable:
echo %PATH%
echo.
echo Press any key to exit...
pause >NUL