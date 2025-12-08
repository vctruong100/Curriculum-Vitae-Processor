@echo off
setlocal enabledelayedexpansion
REM ============================================================================
REM  RUN_ONECLICK_AND_SPLIT_PDF.bat  (uses user's cv_splitter_v2.py with PDF built-in)
REM    1) Run One-Click (No-Year CSV Fix) GUI
REM    2) Split exact FINAL updated CV (explicit path) in Output\ (PDFs created by splitter)
REM ============================================================================

pushd "%~dp0"
set "ROOT=%CD%"
set "PROC=%ROOT%\Processors"
set "OUT=%ROOT%\Output"

if not exist "%OUT%" mkdir "%OUT%"

REM Prefer py launcher; fallback to python
where py >nul 2>nul && set "PY=py" || set "PY=python"

REM Quiet installs needed by GUI + splitter
%PY% -3 -m pip install --upgrade --quiet pip
%PY% -3 -m pip install --quiet python-docx FreeSimpleGUI PySimpleGUI docx2pdf

REM Flags
set "DO_GUI=1"
set "DO_SPLIT=1"
for %%A in (%*) do (
  if /I "%%~A"=="--no-split"   set "DO_SPLIT=0"
  if /I "%%~A"=="--split-only" set "DO_GUI=0"
)

if "%DO_GUI%"=="1" (
  REM --------------------------------------------------------------------------
  REM Step 1: Launch the One-Click (+ No-Year CSV Fix) GUI
  REM --------------------------------------------------------------------------
  %PY% -3 "%PROC%\cv_gui_oneclick_plus_splitter_noyears.py"
  if errorlevel 1 (
    echo GUI returned an error. Aborting.
    popd
    exit /b 1
  )
)

if "%DO_SPLIT%"=="1" (
  REM --------------------------------------------------------------------------
  REM Step 2: Split exact file (explicit path) in Output\ 
  REM   - Find newest non-generated .docx in Output\ and pass it explicitly
  REM     to cv_splitter_v2.py (which will make PDFs itself).
  REM --------------------------------------------------------------------------
  set "FINAL_CV="
  for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command ^
    "(Get-ChildItem -Path '%OUT%' -Filter *.docx | Where-Object { $_.Name -notmatch '^(?i)CenExel( Abbrv)? CURRICULUM VITAE ' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName)"`) do (
      set "FINAL_CV=%%F"
  )

  if "%FINAL_CV%"=="" (
    echo ERROR: Could not find a final updated CV in "%OUT%".
    echo Looked for newest *.docx excluding generated outputs.
    popd
    exit /b 2
  )

  echo Splitting (and converting to PDF): %FINAL_CV%
  pushd "%OUT%"
  set "OUTPUT_DIR=%OUT%"
  set "OUTPUT_DIR=%OUT%"
  %PY% -3 "%PROC%\cv_splitter_v2.py" --outdir "%OUT%" "%FINAL_CV%"
  set "RC=%ERRORLEVEL%"
  popd
  if not "%RC%"=="0" (
    echo Splitter returned an error. Exit code: %RC%
    popd
    exit /b %RC%
  )
)

echo Done.
popd
