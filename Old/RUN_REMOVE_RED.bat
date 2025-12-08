@echo off
setlocal ENABLEDELAYEDEXPANSION
title CV Pipeline — Remove Red Labels

REM Move to script root
pushd "%~dp0"

echo ======================================================
echo   CV Pipeline — Remove Red Labels (Final Pass Only)
echo   This will NOT re-sort or rename any files.
echo ======================================================
echo.

REM ---------- Locate Python ----------
set "PY="

echo Checking for Python via 'py -3' ...
where py >nul 2>nul
if %ERRORLEVEL%==0 (
  py -3 -V >nul 2>nul
  if %ERRORLEVEL%==0 (
    for /f "tokens=*" %%A in ('py -3 -c "import sys;print(sys.executable)"') do set "PY=py -3"
  )
)

if not defined PY (
  echo Checking for Python via 'python' ...
  where python >nul 2>nul
  if %ERRORLEVEL%==0 (
    python -V >nul 2>nul
    if %ERRORLEVEL%==0 set "PY=python"
  )
)

if not defined PY (
  echo Checking for Python via 'python3' ...
  where python3 >nul 2>nul
  if %ERRORLEVEL%==0 (
    python3 -V >nul 2>nul
    if %ERRORLEVEL%==0 set "PY=python3"
  )
)

if not defined PY (
  echo.
  echo ERROR: Python 3 was not found on this system.
  echo Please install Python 3.x and re-run this script.
  echo https://www.python.org/downloads/
  echo.
  pause
  exit /b 1
)

echo Using Python: %PY%
call %PY% -V
echo.

REM ---------- Ensure required packages (NO pip upgrade; no heredocs) ----------
echo Ensuring required Python packages are installed ...
set "PIP=%PY% -m pip"

REM Check pip exists
call %PY% -m pip --version >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
  echo ERROR: pip is not available for %PY%
  echo Try repairing your Python installation.
  echo.
  pause
  exit /b 1
)

REM Check python-docx
call %PY% -c "import docx" >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
  echo Installing python-docx ...
  call %PIP% install python-docx || (echo ERROR installing python-docx & goto :FAIL)
) else (
  echo python-docx already installed.
)

REM Check FreeSimpleGUI, else PySimpleGUI, else install FreeSimpleGUI
call %PY% -c "import FreeSimpleGUI" >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
  call %PY% -c "import PySimpleGUI" >nul 2>nul
  if %ERRORLEVEL% NEQ 0 (
    echo Installing FreeSimpleGUI ...
    call %PIP% install FreeSimpleGUI || (echo ERROR installing FreeSimpleGUI & goto :FAIL)
  ) else (
    echo PySimpleGUI already installed.
  )
) else (
  echo FreeSimpleGUI already installed.
)

REM ---------- Launch the GUI ----------
set "PROC=%~dp0Processors"
if not exist "%PROC%\cv_gui_remove_red.py" (
  echo.
  echo ERROR: "%PROC%\cv_gui_remove_red.py" not found.
  echo Make sure you placed the GUI in the Processors folder.
  echo.
  pause
  exit /b 1
)

echo.
echo Launching GUI ...
echo (If nothing appears, check Windows SmartScreen / antivirus prompts.)
echo.
call %PY% "%PROC%\cv_gui_remove_red.py"
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" (
  echo GUI exited with code %RC%
) else (
  echo Completed.
)
echo.
pause
endlocal
exit /b 0

:FAIL
echo.
echo One or more dependencies failed to install.
echo Please review the errors above.
echo.
pause
endlocal
exit /b 1
