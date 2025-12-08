@echo off
setlocal
pushd "%~dp0"

REM --- Sorting (Phase-aware + MASTER formatting + descending years) ---
REM If your script name differs, update it below.
py -3 "%~dp0sorterv2.py" ^
  --master ".NO_RED_STUDYLIST (EDITABLE).txt" ^
  --unsorted ".ADD_CV_STUDIES (EDITABLE).txt" ^
  --out "SORTED_STUDY_CV_TXT.txt" ^
  --audit "match_audit_report.txt" ^
  --threshold 0.80 ^
  --docx-out "SORTED_STUDY_CV_DOCX.docx" ^
  --indent-type spaces --indent-size 1 ^
  --docx-indent 0.5 ^
  --text-bold-markers false ^
  --bold true

popd
echo.
pause