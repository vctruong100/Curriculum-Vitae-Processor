@echo off
setlocal enabledelayedexpansion
REM Launch the Three-file GUI with built-in splitter
pushd "%~dp0"
REM Expect folder layout:
REM   CV Script\Run_CV_Three_Files_PLUS_SPLITTER.bat
REM   CV Script\Processors\cv_gui_threefile_plus_splitter.py
set ROOT=%CD%
set PROC=%ROOT%\Processors

REM Ensure Python is available (prefer py launcher)
where py >nul 2>nul && set PY=py || set PY=python

REM Create Output folder if missing
if not exist "%ROOT%\Output" mkdir "%ROOT%\Output"

REM Install/upgrade dependencies quietly
%PY% -3 -m pip install --upgrade --quiet pip
%PY% -3 -m pip install --quiet python-docx FreeSimpleGUI PySimpleGUI

REM Run GUI
%PY% -3 "%PROC%\cv_gui_threefile_plus_splitter.py"
popd
