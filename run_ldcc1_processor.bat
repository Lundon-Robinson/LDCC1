@echo off
REM LDCC1 Data Processor Launcher
REM ==============================
REM This batch file launches the LDCC1 Data Processor with proper error handling

echo.
echo LDCC1 Data Processor
echo ===================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.6 or higher from https://python.org
    echo.
    pause
    exit /b 1
)

REM Display Python version
echo Python version:
python --version
echo.

REM Check if required packages are installed
echo Checking dependencies...
python -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing required dependencies...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        echo Please run: pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
)

echo Dependencies OK
echo.

REM Launch the application
echo Starting LDCC1 Data Processor...
echo.
python ldcc1_processor.py

REM Check if the application exited with an error
if errorlevel 1 (
    echo.
    echo Application exited with an error
    pause
) else (
    echo.
    echo Application closed successfully
)