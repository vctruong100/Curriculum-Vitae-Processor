@echo off
setlocal
pushd "%~dp0"
set "ROOT=%CD%"
set "PROC=%ROOT%\Processors"

where py >nul 2>nul && set "PY=py" || set "PY=python"

%PY% -3 -m pip install --upgrade --quiet pip
%PY% -3 -m pip install --quiet FreeSimpleGUI PySimpleGUI python-docx docx2pdf

REM Prefer running from Processors if present
if exist "%PROC%\cv_gui_all_in_one.py" (
  %PY% -3 "%PROC%\cv_gui_all_in_one.py"
) else (
  %PY% -3 "cv_gui_all_in_one.py"
)

popd
