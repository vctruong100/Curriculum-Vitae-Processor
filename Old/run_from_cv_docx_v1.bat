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

REM === CONFIG: name of the full CV .docx placed in the EDITABLE folder ===
set "CV_DOCX=Full CV.docx"

echo.
echo Running CV Processor (DOCX source)...
echo EDITABLE:   "%EDIT%"
echo PROCESSORS: "%PROC%"
echo PLACEHOLDER:"%PLAC%"
echo OUTPUT:     "%OUT%"
echo.
echo Using CV:   "%EDIT%\%CV_DOCX%"
echo.

REM 0) Extract studies from the full CV .docx (Research Experience -> before disclaimer)
py -3 "%PROC%\extract_cv_studies.py" ^
  --cv "%EDIT%\%CV_DOCX%" ^
  --out "%EDIT%\.ADD_CV_STUDIES (EDITABLE).txt"

if errorlevel 1 (
  echo.
  echo [ERROR] Extract step failed. Make sure the CV has a "Research Experience" section
  echo         and the disclaimer line is present. See extractor messages above.
  goto :done
)

REM 1) Sort using MASTER formatting (Phase -> Category, descending years)
py -3 "%PROC%\sorterv2.py" ^
  --master "%EDIT%\.NO_RED_STUDYLIST (EDITABLE).txt" ^
  --unsorted "%EDIT%\.ADD_CV_STUDIES (EDITABLE).txt" ^
  --out "%PLAC%\SORTED_STUDY_CV_TXT.txt" ^
  --audit "%PLAC%\match_audit_report.txt" ^
  --threshold 0.80 ^
  --docx-out "%PLAC%\SORTED_STUDY_CV_DOCX.docx" ^
  --indent-type spaces --indent-size 1 ^
  --docx-indent 0.5 ^
  --text-bold-markers false ^
  --bold true

REM Bail out early if step 1 failed
if errorlevel 1 (
  echo.
  echo [ERROR] Sorting step failed. Check messages above.
  goto :done
)

REM 2) Merge in NEW studies (> latest year), preserving red labels
py -3 "%PROC%\compare_insert_red_docx.py" ^
  --existing-docx "%PLAC%\SORTED_STUDY_CV_DOCX.docx" ^
  --master-docx "%EDIT%\.YES_RED_STUDYLIST (EDITABLE).docx" ^
  --out-docx "%OUT%\.UPDATED CV.docx" ^
  --indent 0.5

:done
popd
echo.
echo Check output file(s) in Output folder.
echo.
pause
