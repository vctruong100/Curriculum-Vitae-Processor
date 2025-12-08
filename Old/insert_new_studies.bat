@echo off
setlocal
pushd "%~dp0"

REM --- Merge: take existing sorted .docx and add NEW master studies (> latest year) ---
REM Keeps Phase -> Category correct and preserves red labels for new inserts.
py -3 "%~dp0compare_insert_red_docx.py" ^
  --existing-docx "SORTED_STUDY_CV_DOCX.docx" ^
  --master-docx ".YES_RED_STUDYLIST (EDITABLE).docx" ^
  --out-docx ".UPDATED CV (FINAL).docx" ^
  --indent 0.5

popd
echo.
pause