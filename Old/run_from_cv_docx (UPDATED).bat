@echo off
setlocal
pushd "%~dp0"

set "ROOT=%~dp0"
set "EDIT=%ROOT%Editable (Add CV here)"
set "PROC=%ROOT%Processors"
set "PLAC=%ROOT%Placeholder"
set "OUT=%ROOT%Output"

if not exist "%EDIT%"  mkdir "%EDIT%"
if not exist "%PROC%"  mkdir "%PROC%"
if not exist "%PLAC%"  mkdir "%PLAC%"
if not exist "%OUT%"   mkdir "%OUT%"

echo.
echo Running CV Processor (DOCX source with auto-detect)...
echo EDITABLE:   "%EDIT%"
echo PROCESSORS: "%PROC%"
echo PLACEHOLDER:"%PLAC%"
echo OUTPUT:     "%OUT%"
echo.

REM === Find the newest CV matching: "CenExel CURRICULUM VITAE Template *.docx"
set "CV_DOCX="
for /f "delims=" %%F in ('dir /b /a:-d /o:-d "%EDIT%\CenExel CURRICULUM VITAE Template *.docx" 2^>nul') do (
  set "CV_DOCX=%%F"
  goto :cvfound
)

echo [ERROR] No CV found matching: "%EDIT%\CenExel CURRICULUM VITAE Template *.docx"
echo         Put the CV in the Editable folder with that naming pattern.
goto :done

:cvfound
echo Using CV: "%EDIT%\%CV_DOCX%"
for %%I in ("%CV_DOCX%") do set "CV_NAME_NOEXT=%%~nI"

echo.
echo [0/4] Extracting studies from detected CV...
py -3 "%PROC%\extract_cv_studies.py" ^
  --cv "%EDIT%\%CV_DOCX%" ^
  --out "%PLAC%\.ADD_CV_STUDIES_FROM_DOCX.txt"

if errorlevel 1 (
  echo.
  echo [ERROR] Extract step failed. Check messages above.
  goto :done
)

echo.
echo "[1/4] Sorting with MASTER (Phase -> Category, desc years)..."
py -3 "%PROC%\sorterv2.py" ^
  --master "%EDIT%\.NO_RED_STUDYLIST (EDITABLE).txt" ^
  --unsorted "%PLAC%\.ADD_CV_STUDIES_FROM_DOCX.txt" ^
  --out "%PLAC%\SORTED_STUDY_CV_TXT.txt" ^
  --audit "%PLAC%\match_audit_report.txt" ^
  --threshold 0.80 ^
  --docx-out "%PLAC%\SORTED_STUDY_CV_DOCX.docx" ^
  --indent-type spaces --indent-size 1 ^
  --docx-indent 0.5 ^
  --text-bold-markers false ^
  --bold true

if errorlevel 1 (
  echo.
  echo [ERROR] Sorting step failed. Check messages above.
  goto :done
)

echo.
echo [2/4] Merging NEW studies from Master .docx, preserving red labels...
py -3 "%PROC%\compare_insert_red_docx.py" ^
  --existing-docx "%PLAC%\SORTED_STUDY_CV_DOCX.docx" ^
  --master-docx "%EDIT%\.YES_RED_STUDYLIST (EDITABLE).docx" ^
  --out-docx "%OUT%\.UPDATED STUDIES.docx" ^
  --indent 0.5

if errorlevel 1 (
  echo.
  echo [WARN] Merge step failed. Proceeding with SORTED_STUDY_CV_DOCX.docx for injection.
  copy /y "%PLAC%\SORTED_STUDY_CV_DOCX.docx" "%OUT%\.UPDATED CV.docx" >nul
)

echo.
echo [3/4] Cloning original CV and injecting sorted studies into Research Experience...
set "OUT_CV=%OUT%\%CV_NAME_NOEXT% (Updated).docx"
py -3 "%PROC%\inject_sorted_into_cv.py" ^
  --original-cv "%EDIT%\%CV_DOCX%" ^
  --studies-docx "%OUT%\.UPDATED STUDIES.docx" ^
  --out "%OUT_CV%"

if errorlevel 1 (
  echo.
  echo [ERROR] Injection step failed. Check messages above.
  goto :done
)

echo.
echo [4/4] All done.
echo Output files:
echo   - "%OUT%\.UPDATED CV.docx"   (studies document)
echo   - "%OUT_CV%"                 (full CV with injected studies)

:done
popd
echo.
echo Check output files in the Output folder.
echo.
pause
