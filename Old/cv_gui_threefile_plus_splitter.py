#!/usr/bin/env python3
# Three-file CV Pipeline GUI (+ Splitter)
# Inputs:
#   1) Original CV (.docx)
#   2) Unred MASTER schedule (.txt)  -> for sorting
#   3) MASTER with red labels (.docx) -> for merge (preserve red)
#
# Outputs:
#   - Final injected CV to: CV Script\Output\<same name as original>.docx
#   - (Optional) Split outputs (from the final CV) to:
#       * CenExel CURRICULUM VITAE {Name}.docx
#       * CenExel Abbrv CURRICULUM VITAE {Name}.docx

import os, sys, subprocess
from pathlib import Path

try:
    import FreeSimpleGUI as sg
except Exception:
    try:
        import PySimpleGUI as sg
    except Exception:
        raise SystemExit("FreeSimpleGUI not installed. Run via the .bat launcher to auto-install.")

HERE = Path(__file__).resolve().parent   # ...\CV Script\Processors
ROOT = HERE.parent                        # ...\CV Script
OUT  = ROOT / "Output"
OUT.mkdir(parents=True, exist_ok=True)

# Processor scripts (must live in the same folder as this GUI)
SCRIPT_EXTRACT = HERE / "extract_cv_studies.py"
SCRIPT_SORT    = HERE / "sorterv2.py"
SCRIPT_MERGE   = HERE / "compare_insert_red_docx.py"
SCRIPT_INJECT  = HERE / "inject_sorted_into_cv.py"
SCRIPT_SPLIT   = HERE / "cv_splitter_v2.py"

SECTION_START  = "Research Experience"
SECTION_END    = "By signing this form, I confirm that the information provided is accurate and reflects my current qualifications."

# Intermediates (deleted after success)
UNSORTED_TXT   = OUT / ".ADD_CV_STUDIES_FROM_DOCX.txt"
SORTED_TXT     = OUT / "SORTED_STUDY_CV_TXT.txt"
AUDIT_TSV      = OUT / "match_audit_report.tsv"
SORTED_DOCX    = OUT / "SORTED_STUDY_CV_DOCX.docx"
MERGED_DOCX    = OUT / ".UPDATED CV.docx"

def exe():
    return sys.executable or "python"

def norm(p: Path) -> str:
    return str(p.expanduser().resolve())

def run_cmd(window, args, cwd=None):
    window["-LOG-"].print(f"$ {' '.join(args)}", text_color="yellow")
    try:
        p = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, cwd=cwd)
        for line in p.stdout:
            window["-LOG-"].print(line.rstrip())
        rc = p.wait()
        if rc != 0:
            window["-LOG-"].print(f"Exited with code {rc}", text_color="red")
            sg.popup_error("A step failed. See log for details.")
            return False
        window["-LOG-"].print("Done.", text_color="green")
        return True
    except Exception as e:
        window["-LOG-"].print(f"ERROR: {e}", text_color="red")
        sg.popup_error(f"ERROR: {e}")
        return False

def validate_inputs(cv_path: Path, master_txt: Path, master_red_docx: Path):
    if not cv_path.is_file() or cv_path.suffix.lower() != ".docx":
        return "Select a valid ORIGINAL CV .docx"
    if not master_txt.is_file() or master_txt.suffix.lower() != ".txt":
        return "Select a valid UNRED MASTER schedule .txt"
    if not master_red_docx.is_file() or master_red_docx.suffix.lower() != ".docx":
        return "Select a valid MASTER schedule with red labels .docx"
    return None

# ---------- UI ----------
sg.theme("SystemDefault")
layout = [
    [sg.Text("Select the three inputs, toggle Split if desired, then press Process →", font=("Segoe UI", 11))],
    [sg.Text("Original CV (.docx)", size=(22,1)), sg.Input(key="-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
    [sg.Text("Unred MASTER (.txt)", size=(22,1)), sg.Input(key="-MASTER_TXT-", expand_x=True), sg.FileBrowse(file_types=(("Text", "*.txt"),))],
    [sg.Text("MASTER w/ red (.docx)", size=(22,1)), sg.Input(key="-MASTER_RED-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
    [sg.Checkbox("Also create Abbreviated & Full CV from the FINAL CV (run splitter)", key="-DO_SPLIT-", default=True)],
    [sg.Button("Process", bind_return_key=True, size=(18,1)), sg.Button("Open Output"), sg.Button("Quit")],
    [sg.Frame("Log", [[sg.Multiline(size=(100,22), key="-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)]
]
window = sg.Window("CV Pipeline — Three Files (+ Splitter)", layout, finalize=True)

def process_all(window, cv_file: str, master_txt_file: str, master_red_file: str, do_split: bool):
    cv_path          = Path(cv_file.strip())
    master_txt       = Path(master_txt_file.strip())
    master_red_docx  = Path(master_red_file.strip())

    err = validate_inputs(cv_path, master_txt, master_red_docx)
    if err:
        sg.popup_error(err)
        return

    final_cv = OUT / cv_path.name  # keep the same filename as the original

    window["-LOG-"].print(f"Original CV: {cv_path}")
    window["-LOG-"].print(f"Unred MASTER (.txt): {master_txt}")
    window["-LOG-"].print(f"MASTER with red (.docx): {master_red_docx}")

    # 1) Extract studies from CV to UNSORTED_TXT
    args = [exe(), norm(SCRIPT_EXTRACT),
            "--cv", norm(cv_path),
            "--out", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(window, args): return

    # 2) Sort using unred master .txt  -> produce SORTED_DOCX (and intermediates)
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

    # 3) Merge to preserve red labels using the red-labeled master .docx
    args = [exe(), norm(SCRIPT_MERGE),
            "--existing-docx", norm(SORTED_DOCX),
            "--master-docx",   norm(master_red_docx),
            "--out-docx",      norm(MERGED_DOCX),
            "--indent",        "0.5"]
    if not run_cmd(window, args): return

    # 4) Inject merged studies into original CV -> final file (same name as original)
    args = [exe(), norm(SCRIPT_INJECT),
            "--original-cv",  norm(cv_path),
            "--studies-docx", norm(MERGED_DOCX),
            "--out",          norm(final_cv),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(window, args): return

    # 5) Optional: Split the FINAL CV into Abbreviated + Full
    if do_split:
        window["-LOG-"].print("Running splitter on FINAL CV...")
        args = [exe(), norm(SCRIPT_SPLIT), norm(final_cv)]
        if not run_cmd(window, args, cwd=OUT): return

    # Clean intermediates so only the final/split files remain
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
        process_all(window, val["-CV-"], val["-MASTER_TXT-"], val["-MASTER_RED-"], bool(val["-DO_SPLIT-"]))

window.close()
