@echo off
setlocal
pushd "%~dp0"

REM Root folders
set "ROOT=%CD%"
set "PROC=%ROOT%\Processors"
set "EDIT=%ROOT%\Editable"

REM Prefer 'py' if available, otherwise 'python'
where py >nul 2>nul && set "PY=py" || set "PY=python"

echo ---------------------------------------------------------
echo   Generating .NO_RED_STUDYLIST (EDITABLE).txt from CSV
echo ---------------------------------------------------------

REM ----------------------------------------
REM Check that the converter script exists
REM ----------------------------------------
if not exist "%PROC%\csv_to_no_red_master.py" (
  echo ERROR: Missing script:
  echo   %PROC%\csv_to_no_red_master.py
  echo Make sure the file exists.
  pause
  goto :EOF
)

REM ----------------------------------------
REM Check that MASTER CSV exists
REM ----------------------------------------
if not exist "%EDIT%\Master study list.csv" (
  echo ERROR: Master study list CSV not found:
  echo   %EDIT%\Master study list.csv
  echo Fix the name or move the file into Editable folder.
  pause
  goto :EOF
)

REM ----------------------------------------
REM RUN THE SCRIPT
REM ----------------------------------------
%PY% -3 "%PROC%\csv_to_no_red_master.py" ^
  --csv "%EDIT%\Master study list.csv" ^
  --out "%EDIT%\.NO_RED_STUDYLIST (EDITABLE).txt" ^
  --has-header

if errorlevel 1 (
  echo.
  echo ERROR: Conversion script failed.
  pause
  goto :EOF
)

echo.
echo SUCCESS!
echo Updated: "%EDIT%\.NO_RED_STUDYLIST (EDITABLE).txt"
echo Source:  "%EDIT%\Master study list.csv"
echo.

popd
endlocal
