@echo off
setlocal
pushd "%~dp0"

echo ==============================================
echo   CV Processor - One Click
echo ==============================================
echo.

set "LOG=%~dp0last_run.log"
echo ==== %date% %time% ==== > "%LOG%"

where py >nul 2>nul || (echo ERROR: Python 3.8+ not found. & pause & exit /b 1)

REM Install free GUI lib + docx if missing
py -3 -m pip show FreeSimpleGUI >nul 2>nul || py -3 -m pip install --quiet FreeSimpleGUI
py -3 -m pip show python-docx   >nul 2>nul || py -3 -m pip install --quiet python-docx

REM Ensure Output folder exists
if not exist "%~dp0Output" mkdir "%~dp0Output"

REM Run the one-click GUI from Processors
cd /d "%~dp0Processors"
echo Starting GUI...
echo.
py -3 "%~dp0Processors\cv_gui.py" 1>>"%LOG%" 2>&1

echo.
echo ===== Console output (this run) =====
type "%LOG%"
echo =====================================
echo (Full log saved to: "%LOG%")
echo.
pause
