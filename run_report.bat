@echo off
REM =====================================================
REM Large Deal Report Automation - Windows 11
REM Double-click this file to run the daily report
REM =====================================================

echo.
echo ========================================
echo  Large Deal Report Automation
echo ========================================
echo.

cd /d "%~dp0"
python large_deal_report_automation.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================
    echo  ERROR: Automation failed!
    echo  Exit code: %ERRORLEVEL%
    echo ========================================
    echo.
    echo Press any key to close...
    pause >nul
) else (
    echo.
    echo Press any key to close...
    pause >nul
)
