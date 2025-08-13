@echo off
REM This batch file installs the required Python libraries for the AutoPrint application.

REM Check if Python is installed
python --version >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python and try again.
    pause
    exit /b 1
)

REM Check if pip is installed
pip --version >nul 2>nul
if %errorlevel% neq 0 (
    echo pip is not installed. Please ensure you have a full Python installation with pip.
    pause
    exit /b 1
)

echo Installing required Python libraries...

pip install Pillow pystray

if %errorlevel% equ 0 (
    echo Dependencies installed successfully.
) else (
    echo There was an error during installation.
)

pause