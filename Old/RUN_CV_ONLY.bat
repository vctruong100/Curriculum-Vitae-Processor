@echo off
setlocal enabledelayedexpansion
REM Launch the One-Click GUI with built-in splitter
pushd "%~dp0"
REM Expect folder layout:
REM   CV Script\Run_CV_GUI_ONECLICK_PLUS_SPLITTER.bat
REM   CV Script\Processors\cv_gui_oneclick_plus_splitter.py
set ROOT=%CD%
set PROC=%ROOT%\Processors

where py >nul 2>nul && set PY=py || set PY=python

if not exist "%ROOT%\Output" mkdir "%ROOT%\Output"
if not exist "%ROOT%\Editable" mkdir "%ROOT%\Editable"

%PY% -3 -m pip install --upgrade --quiet pip
%PY% -3 -m pip install --quiet python-docx FreeSimpleGUI PySimpleGUI

%PY% -3 "%PROC%\cv_gui_oneclick_plus_splitter.py"
popd
