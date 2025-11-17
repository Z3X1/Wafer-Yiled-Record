@echo off
REM Wafer Yield Analyzer - Windows Batch Script
REM This script runs the wafer yield analysis

echo ========================================
echo Wafer Yield Analyzer
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.10 or higher
    pause
    exit /b 1
)

echo Checking dependencies...
pip show pandas >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
)

echo.
echo Running analysis...
echo ========================================
python wafer_yield_analyzer.py
if errorlevel 1 (
    echo.
    echo ERROR: Analysis failed
    pause
    exit /b 1
)

echo.
echo ========================================
echo Analysis completed successfully!
echo.
echo Output files:
echo   - wafer_yield_report.xlsx
echo   - wafer_yield_chart.png
echo ========================================
echo.

pause

