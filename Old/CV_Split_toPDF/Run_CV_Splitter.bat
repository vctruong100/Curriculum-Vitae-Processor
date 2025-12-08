@echo off
setlocal
pushd "%~dp0"

REM ==========================================================
REM  Run_CV_Splitter.bat  (CLEAN VERSION)
REM  - Detects Python via launcher (py.exe) or python/python3
REM  - Ensures pip + required libs (python-docx, docx2pdf)
REM  - Runs cv_splitter_v2.py with NO arguments (auto-detects input)
REM  NOTE: Do not paste Python code into this .bat file.
REM ==========================================================

REM Prefer Python launcher
set "PY_EXE=%LOCALAPPDATA%\Programs\Python\Launcher\py.exe"
set "PY_FLAG=-3"

if exist "%PY_EXE%" (
  echo Using: "%PY_EXE%" %PY_FLAG%
) else (
  set "PY_EXE=python"
  set "PY_FLAG="
  echo Trying fallback: %PY_EXE%
)

REM Verify Python runs
"%PY_EXE%" %PY_FLAG% -V
if errorlevel 1 (
  echo [ERROR] Could not run Python.
  echo If Windows opens Microsoft Store, disable these aliases:
  echo   Settings > Apps > Advanced app settings > App execution aliases
  echo Turn OFF "python.exe" and "python3.exe".
  pause
  exit /b 1
)

REM Ensure pip is present and up to date
"%PY_EXE%" %PY_FLAG% -m ensurepip --upgrade
"%PY_EXE%" %PY_FLAG% -m pip install --upgrade pip

REM Install required libraries
echo Installing required libraries...
"%PY_EXE%" %PY_FLAG% -m pip install python-docx docx2pdf

echo.
echo Running CV Splitter Script...
echo.

REM Run the Python script exactly as-is (no args; it will auto-find the docx)
"%PY_EXE%" %PY_FLAG% "%~dp0cv_splitter_v2.py"
set "ERR=%ERRORLEVEL%"

if not "%ERR%"=="0" (
  echo.
  echo Script failed. Exit code: %ERR%
  pause
  exit /b %ERR%
)

echo.
echo Done! PDFs saved next to your input .docx.
pause

popd
endlocal
