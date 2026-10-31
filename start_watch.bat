@echo off
title Credit Recovery - data watch
cd /d "%~dp0.."
echo ============================================================
echo  Credit Recovery dashboard - automatic data refresh
echo ============================================================
echo  Keeping this window open regenerates data.js whenever a
echo  workbook is added or replaced in:
echo    - Daily Overdue Data
echo    - Daily Consolidate for Overdue Report
echo    - NPL Movement
echo    - Data Dashboard.xlsx
echo.
echo  Leave it running, then keep index.html open in the browser.
echo  Close this window (or press Ctrl+C) to stop watching.
echo ============================================================
echo.
python automation\build_data.py --watch
echo.
echo Watch stopped.
pause
