#!/usr/bin/env python3
# One-Click CV Pipeline GUI (+ Splitter)
# Drop original CV .docx, uses MASTER files from "Editable" if present.
# Outputs to: CV Script\Output\

import os, sys, subprocess, glob
from pathlib import Path

try:
    import FreeSimpleGUI as sg
except Exception:
    try:
        import PySimpleGUI as sg
    except Exception:
        raise SystemExit("FreeSimpleGUI not installed. Run via the .bat launcher to auto-install.")

HERE = Path(__file__).resolve().parent         # ...\CV Script\Processors
ROOT = HERE.parent                              # ...\CV Script
EDIT = ROOT / "Editable"
OUT  = ROOT / "Output"
OUT.mkdir(parents=True, exist_ok=True)

# Processor scripts
SCRIPT_EXTRACT = HERE / "extract_cv_studies.py"
SCRIPT_SORT    = HERE / "sorterv2.py"
SCRIPT_MERGE   = HERE / "compare_insert_red_docx.py"
SCRIPT_INJECT  = HERE / "inject_sorted_into_cv.py"
SCRIPT_SPLIT   = HERE / "cv_splitter_v2.py"

SECTION_START  = "Research Experience"
SECTION_END    = "By signing this form, I confirm that the information provided is accurate and reflects my current qualifications."

# Fixed output paths
UNSORTED_TXT   = OUT / ".ADD_CV_STUDIES_FROM_DOCX.txt"
SORTED_TXT     = OUT / "SORTED_STUDY_CV_TXT.txt"
AUDIT_TSV      = OUT / "match_audit_report.tsv"
SORTED_DOCX    = OUT / "SORTED_STUDY_CV_DOCX.docx"
MERGED_DOCX    = OUT / ".UPDATED CV.docx"           

def exe():
    return sys.executable or "python"

def norm(p):
    return str(Path(p).expanduser().resolve())

def run_cmd(window, args, cwd=None):
    window["-LOG-"].print(f"$ {' '.join(args)}", text_color="yellow")
    try:
        p = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, cwd=cwd)
        for line in p.stdout:
            window["-LOG-"].print(line.rstrip())
        rc = p.wait()
        if rc != 0:
            window["-LOG-"].print(f"Exited with code {rc}", text_color="red")
            sg.popup_error("A step failed. Check the log.")
            return False
        window["-LOG-"].print("Done.", text_color="green")
        return True
    except Exception as e:
        window["-LOG-"].print(f"ERROR: {e}", text_color="red")
        sg.popup_error(f"ERROR: {e}")
        return False

def find_master_txt():
    candidates = [
        EDIT / ".NO_RED_STUDYLIST (EDITABLE).txt",
        EDIT / "NO_RED_STUDYLIST (EDITABLE).txt",
        EDIT / ".YES_RED_STUDYLIST (EDITABLE).docx",
        ROOT / ".NO_RED_STUDYLIST (EDITABLE).txt",
        ROOT / "NO_RED_STUDYLIST (EDITABLE).txt",
        ROOT / ".YES_RED_STUDYLIST (EDITABLE).docx",
    ]
    for c in candidates:
        if "*" in str(c):
            matches = sorted([Path(p) for p in glob.glob(str(c))])
            if matches:
                return matches[0]
        elif c.is_file():
            return c
    return None

def find_master_docx():
    candidates = [
        EDIT / ".UPDATED CV.docx",
        EDIT / "UPDATED CV.docx",
        EDIT / ".YES_RED_STUDYLIST (EDITABLE).docx",
        ROOT / ".UPDATED CV.docx",
        ROOT / "UPDATED CV.docx",
        ROOT / ".YES_RED_STUDYLIST (EDITABLE).docx",
    ]
    for c in candidates:
        if "*" in str(c):
            matches = sorted([Path(p) for p in glob.glob(str(c))])
            if matches:
                return matches[0]
        elif c.is_file():
            return c
    return None

sg.theme("SystemDefault")

layout = [
    [sg.Text("Browse and Select the ORIGINAL CV (.docx). Toggle Split if desired, then Process →", font=("Segoe UI", 11))],
    [sg.Input(key="-CV-", expand_x=True, enable_events=True, tooltip="Drag & drop a .docx here"),
     sg.FileBrowse("Browse", file_types=(("Word", "*.docx"),))],
    [sg.Checkbox("Also create Abbreviated & Full CV from the FINAL CV (run splitter)", key="-DO_SPLIT-", default=True)],
    [sg.Button("Process", bind_return_key=True, size=(18,1)), sg.Button("Open Output"), sg.Button("Quit")],
    [sg.Frame("Log", [[sg.Multiline(size=(100,22), key="-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)]
]
window = sg.Window("CV Pipeline — One Click (+ Splitter)", layout, finalize=True)

def process_all(window, cv_path, do_split: bool):
    cv_path = Path(cv_path)
    final_cv = OUT / cv_path.name

    if not cv_path.is_file():
        sg.popup_error("Please provide the ORIGINAL CV .docx")
        return

    master_txt = find_master_txt()
    if not master_txt:
        sg.popup_error("MASTER .txt not found. Please place a MASTER study list inside 'Editable\\' (e.g., '.NO_RED_STUDYLIST (EDITABLE).txt').")
        return
    window["-LOG-"].print(f"Using MASTER .txt: {master_txt}")

    # 1) Extract
    args = [exe(), norm(SCRIPT_EXTRACT),
            "--cv", norm(cv_path),
            "--out", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end", SECTION_END]
    if not run_cmd(window, args): return

    # 2) Sort (always produce docx)
    args = [exe(), norm(SCRIPT_SORT),
            "--master",        norm(master_txt),
            "--unsorted",      norm(UNSORTED_TXT),
            "--out",           norm(SORTED_TXT),
            "--audit",         norm(AUDIT_TSV),
            "--threshold",     "0.80",
            "--docx-out",      norm(SORTED_DOCX),
            "--docx-indent",   "0.5",
            "--indent-type",   "spaces",
            "--indent-size",   "1",
            "--text-bold-markers", "false",
            "--bold",          "true"]
    if not run_cmd(window, args): return

    # 3) Merge (optional) to preserve red labels, if a MASTER .docx exists
    master_docx = find_master_docx()
    use_for_inject = SORTED_DOCX
    if master_docx and master_docx.is_file():
        window["-LOG-"].print(f"MASTER .docx found: {master_docx} → merging to preserve red labels")
        args = [exe(), norm(SCRIPT_MERGE),
                "--existing-docx", norm(SORTED_DOCX),
                "--master-docx",   norm(master_docx),
                "--out-docx",      norm(MERGED_DOCX),
                "--indent",        "0.5"]
        if not run_cmd(window, args): return
        use_for_inject = MERGED_DOCX
    else:
        window["-LOG-"].print("No MASTER .docx found — skipping merge step.")

    # 4) Inject into original CV
    args = [exe(), norm(SCRIPT_INJECT),
            "--original-cv",  norm(cv_path),
            "--studies-docx", norm(use_for_inject),
            "--out",          norm(final_cv),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(window, args): return

    # 5) Optional: Split the FINAL CV into Abbreviated + Full
    if do_split:
        window["-LOG-"].print("Running splitter on FINAL CV...")
        args = [exe(), norm(SCRIPT_SPLIT), norm(final_cv)]
        if not run_cmd(window, args, cwd=OUT): return

    for p in [UNSORTED_TXT, SORTED_TXT, AUDIT_TSV, SORTED_DOCX, MERGED_DOCX]:
        try:
            if p.exists():
                p.unlink()
        except Exception:
            pass

    window["-LOG-"].print("")
    window["-LOG-"].print(f"✓ Done. Output folder:", text_color="green")
    window["-LOG-"].print(f" - {OUT}")

while True:
    ev, val = window.read()
    if ev in (sg.WIN_CLOSED, "Quit"):
        break
    if ev == "Open Output":
        try:
            os.startfile(str(OUT))
        except Exception:
            subprocess.run(["explorer", str(OUT)])
    if ev == "Process":
        process_all(window, val["-CV-"].strip(), bool(val["-DO_SPLIT-"]))

window.close()
