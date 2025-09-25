#!/bin/bash
# LDCC1 Data Processor Launcher
# ==============================
# This script launches the LDCC1 Data Processor with proper error handling

echo
echo "LDCC1 Data Processor"
echo "==================="
echo

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo "ERROR: Python is not installed"
        echo "Please install Python 3.6 or higher"
        exit 1
    else
        PYTHON_CMD="python"
    fi
else
    PYTHON_CMD="python3"
fi

# Display Python version
echo "Python version:"
$PYTHON_CMD --version
echo

# Check Python version
PYTHON_VERSION=$($PYTHON_CMD -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
REQUIRED_VERSION="3.6"

if [ "$(printf '%s\n' "$REQUIRED_VERSION" "$PYTHON_VERSION" | sort -V | head -n1)" != "$REQUIRED_VERSION" ]; then
    echo "ERROR: Python version $PYTHON_VERSION is not supported"
    echo "Please install Python $REQUIRED_VERSION or higher"
    exit 1
fi

# Check if required packages are installed
echo "Checking dependencies..."
$PYTHON_CMD -c "import pandas, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing required dependencies..."
    pip3 install -r requirements.txt || pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to install dependencies"
        echo "Please run: pip install -r requirements.txt"
        exit 1
    fi
fi

echo "Dependencies OK"
echo

# Launch the application
echo "Starting LDCC1 Data Processor..."
echo
$PYTHON_CMD ldcc1_processor.py

# Check exit status
if [ $? -eq 0 ]; then
    echo
    echo "Application closed successfully"
else
    echo
    echo "Application exited with an error"
    read -p "Press Enter to continue..."
fi