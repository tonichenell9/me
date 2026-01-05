@echo off
REM Large Deal Report Automation - Windows Batch Script
REM Run this script to execute the automation

cd /d "%~dp0"
python large_deal_report_automation.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Error: Automation failed with exit code %ERRORLEVEL%
    pause
)

