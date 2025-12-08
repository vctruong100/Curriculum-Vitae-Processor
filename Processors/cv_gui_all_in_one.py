#!/usr/bin/env python3
# CV Pipeline — All-in-One GUI (3 Tabs)
# Tabs:
#   1) One Click (+ Splitter) with CSV No-Year Fix
#   2) Remove Red Labels (CSV Fuzzy) + Optional Splitter (NEW)
#   3) Three Files (+ Splitter)

import os, sys, subprocess
from pathlib import Path

try:
    import FreeSimpleGUI as sg
except Exception:
    try:
        import PySimpleGUI as sg
    except Exception:
        raise SystemExit("FreeSimpleGUI not installed. Run via the .bat launcher to auto-install.")

HERE = Path(__file__).resolve().parent      # Processors/ or where this file lives
ROOT = HERE.parent if (HERE.name.lower() == "processors") else HERE
OUT  = ROOT / "Output"
OUT.mkdir(parents=True, exist_ok=True)

# Processor scripts
SCRIPT_EXTRACT = HERE / "extract_cv_studies.py"
SCRIPT_RESOLVE = HERE / "resolve_noyear_from_csv.py"
SCRIPT_SORT    = HERE / "sorterv2.py"
SCRIPT_MERGE   = HERE / "compare_insert_red_docx.py"
SCRIPT_INJECT  = HERE / "inject_sorted_into_cv.py"
SCRIPT_SPLIT   = HERE / "cv_splitter_v2.py"
SCRIPT_REMOVE  = HERE / "remove_red_labels_from_docx.py"
SCRIPT_CSV2MASTER = HERE / "csv_to_no_red_master.py"

SECTION_START  = "Research Experience"
SECTION_END    = "By signing this form, I confirm that the infor...on provided is accurate and reflects my current qualifications."

UNSORTED_TXT   = OUT / ".ADD_CV_STUDIES_FROM_DOCX.txt"
SORTED_TXT     = OUT / "SORTED_STUDY_CV_TXT.txt"
AUDIT_TSV      = OUT / "match_audit_report.tsv"
NOYEAR_AUDIT   = OUT / "noyear_resolve_audit.tsv"
SORTED_DOCX    = OUT / "SORTED_STUDY_CV_DOCX.docx"
MERGED_DOCX    = OUT / ".UPDATED CV.docx"

MASTER_TXT     = ROOT / "Editable" / ".NO_RED_STUDYLIST (EDITABLE).txt"
MASTER_TXT_B   = ROOT / "Editable" / ".NO_RED_STUDYLIST_COLB (TEMP).txt"

def exe():
    return sys.executable or "python"

def norm(p: Path) -> str:
    return str(p.expanduser().resolve())

def run_cmd(window, args, cwd=None):
    window.print(f"$ {' '.join(args)}", text_color="yellow")
    try:
        p = subprocess.Popen(args, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    except Exception as e:
        window.print(f"ERROR: {e}", text_color="red")
        return False
    for line in p.stdout:
        window.print(line.rstrip("\n"))
    p.wait()
    if p.returncode != 0:
        window.print(f"[EXIT CODE {p.returncode}]", text_color="red")
        return False
    return True

def move_split_outputs_to_out(final_cv: Path, outdir: Path, logwin):
    base = final_cv.stem
    parent = final_cv.parent
    candidates = [
        parent / f"{base} (Abbreviated).docx",
        parent / f"{base} (Full).docx",
        parent / f"{base} (Abbreviated CV).docx",
        parent / f"{base} (Full CV).docx",
    ]
    for c in candidates:
        if c.is_file():
            dest = outdir / c.name
            try:
                if dest.exists():
                    dest.unlink()
                c.replace(dest)
                logwin.print(f"Moved split output: {c} → {dest}")
            except Exception as e:
                logwin.print(f"Could not move split output {c}: {e}", text_color="red")

# ---------------- Tab 1: One-Click CV Pipeline ----------------
def tab1_layout():
    return [
        [sg.Text("Original CV (.docx)", size=(22,1)), sg.Input(key="-T1-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Text("Master CSV (Phase/Category/Year, Red, No-Red)", size=(22,2)), sg.Input(key="-T1-CSV-", expand_x=True), sg.FileBrowse(file_types=(("CSV", "*.csv"),))],
        [sg.Text("No-Year resolve threshold (0.80–0.98)"), sg.Slider(range=(0.80,0.98), resolution=0.01, default_value=0.88, orientation="h", size=(30,20), key="-T1-THRESH-")],
        [sg.Checkbox("Also create Abbreviated & Full CV from the UPDATED CV (run splitter)", key="-T1-SPLIT-", default=True)],
        [sg.Button("Run All-in-One", key="-T1-RUN-", size=(18,1)), sg.Button("Open Output", key="-T1-OPEN-")],
        [sg.Frame("Log", [[sg.Multiline(size=(100,18), key="-T1-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)],
    ]

def tab2_layout():
    return [
        [sg.Text("Final CV (.docx)", size=(22,1)), sg.Input(key="-T2-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Text("Mapping CSV", size=(22,1)), sg.Input(key="-T2-CSV-", expand_x=True), sg.FileBrowse(file_types=(("CSV", "*.csv"),))],
        [sg.Text("Fuzzy threshold (0.80–0.98)"), sg.Slider(range=(0.80,0.98), resolution=0.01, default_value=0.90, orientation="h", size=(30,20), key="-T2-TH-")],
        [sg.Checkbox("Also create Abbreviated & Full CV from the UPDATED CV (run splitter)", key="-T2-SPLIT-", default=True)],
        [sg.Button("Run Remove-Red", key="-T2-RUN-", size=(18,1)), sg.Button("Open Output", key="-T2-OPEN-")],
        [sg.Frame("Log", [[sg.Multiline(size=(100,18), key="-T2-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)],
    ]

def tab3_layout():
    return [
        [sg.Text("Original CV (.docx)", size=(22,1)), sg.Input(key="-T3-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Text("Master Red-Label CV (.docx)", size=(22,1)), sg.Input(key="-T3-MASTER-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Text("Sorted No-Red DOCX (from sorter)", size=(22,1)), sg.Input(key="-T3-SORTED-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Checkbox("Also create Abbreviated & Full CV from the UPDATED CV (run splitter)", key="-T3-SPLIT-", default=True)],
        [sg.Button("Run Merge+Inject", key="-T3-RUN-", size=(18,1)), sg.Button("Open Output", key="-T3-OPEN-")],
        [sg.Frame("Log", [[sg.Multiline(size=(100,18), key="-T3-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)],
    ]

def tab1_process(values, logwin):
    cv = Path((values.get("-T1-CV-") or "").strip())
    csvp = Path((values.get("-T1-CSV-") or "").strip())
    try:
        th = float((values.get("-T1-THRESH-") or "0.88").strip())
    except Exception:
        th = 0.88
    do_split = bool(values.get("-T1-SPLIT-"))

    if not cv.is_file() or cv.suffix.lower() != ".docx":
        sg.popup_error("Select a valid ORIGINAL CV (.docx)")
        return
    if not csvp.is_file() or csvp.suffix.lower() != ".csv":
        sg.popup_error("Select a valid mapping CSV (.csv)")
        return

    final_cv = OUT / cv.name

    args = [exe(), norm(SCRIPT_EXTRACT),
            "--cv", norm(cv),
            "--out", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(logwin, args):
        return

    args = [exe(), norm(SCRIPT_RESOLVE),
            "--cv", norm(cv),
            "--csv", norm(csvp),
            "--in-unsorted", norm(UNSORTED_TXT),
            "--out-unsorted", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END,
            "--threshold",     str(th),
            "--audit",         norm(NOYEAR_AUDIT)]
    if not run_cmd(logwin, args):
        return

    args = [exe(), norm(SCRIPT_CSV2MASTER),
            "--csv",      norm(csvp),
            "--out",      norm(MASTER_TXT),
            "--out-b",    norm(MASTER_TXT_B),
            "--has-header"]
    if not run_cmd(logwin, args):
        return

    args = [exe(), norm(SCRIPT_SORT),
        "--master",        norm(MASTER_TXT),
        "--master-b",      norm(MASTER_TXT_B),
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
    if not run_cmd(logwin, args):
        return

    try:
        if MASTER_TXT.exists():
            MASTER_TXT.unlink()
    except Exception as e:
        logwin.print(f"Could not delete {MASTER_TXT}: {e}", text_color="red")

    # 4) Merge if MASTER red docx exists
    master_red_candidates = [
        Path("Editable") / ".UPDATED CV.docx",
        Path("Editable") / "UPDATED CV.docx",
        Path("Editable") / ".YES_RED_STUDYLIST (EDITABLE).docx",
        Path(".UPDATED CV.docx"),
        Path("UPDATED CV.docx"),
        Path(".YES_RED_STUDYLIST (EDITABLE).docx"),
    ]
    use_for_inject = SORTED_DOCX
    for c in master_red_candidates:
        if c.is_file():
            logwin.print(f"MASTER .docx found: {c} → merging to preserve red labels")
            args = [exe(), norm(SCRIPT_MERGE),
                    "--existing-docx", norm(SORTED_DOCX),
                    "--master-docx",   norm(c),
                    "--out-docx",      norm(MERGED_DOCX),
                    "--indent",        "0.5"]
            if not run_cmd(logwin, args):
                return
            use_for_inject = MERGED_DOCX
            break

    # 5) Inject
    args = [exe(), norm(SCRIPT_INJECT),
            "--original-cv",   norm(cv),
            "--studies-docx",  norm(use_for_inject),
            "--out",           norm(final_cv),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(logwin, args):
        return

    # 6) Optional splitter
    if do_split:
        args = [exe(), norm(SCRIPT_SPLIT), "--outdir", norm(OUT), norm(final_cv)]
        if not run_cmd(logwin, args):
            return
        move_split_outputs_to_out(final_cv, OUT, logwin)

    # CLEAN UP TEMPORARY PIPELINE FILES
    temp_files = [
        UNSORTED_TXT,
        NOYEAR_AUDIT,
        SORTED_TXT,
        SORTED_DOCX,
        MERGED_DOCX
    ]

    for p in temp_files:
        try:
            if p.exists():
                p.unlink()
        except Exception as e:
            logwin.print(f"Warning: Could not delete temporary file {p}: {e}", text_color="yellow")


    logwin.print(f"Success! Final CV: {final_cv}", text_color="green")

# ---------------- Tab 2: Remove Red Labels (CSV Fuzzy) + Splitter ----------------
def tab2_process(values, logwin):
    cv = Path((values.get("-T2-CV-") or "").strip())
    csvp = Path((values.get("-T2-CSV-") or "").strip())
    try:
        th = float((values.get("-T2-TH-") or "0.90").strip())
    except Exception:
        th = 0.90
    do_split = bool(values.get("-T2-SPLIT-"))

    if not cv.is_file() or cv.suffix.lower() != ".docx":
        sg.popup_error("Select a valid UPDATED CV (.docx)")
        return
    if not csvp.is_file() or csvp.suffix.lower() != ".csv":
        sg.popup_error("Select a valid mapping CSV (.csv)")
        return

    updated_cv = cv
    cleaned_cv = OUT / (cv.stem + " (No Red Labels).docx")

    args = [exe(), norm(SCRIPT_REMOVE),
            "--cv",      norm(updated_cv),
            "--csv",     norm(csvp),
            "--out",     norm(cleaned_cv),
            "--threshold", str(th)]
    if not run_cmd(logwin, args):
        return

    if do_split:
        args = [exe(), norm(SCRIPT_SPLIT), "--outdir", norm(OUT), norm(cleaned_cv)]
        if not run_cmd(logwin, args):
            return
        move_split_outputs_to_out(cleaned_cv, OUT, logwin)

    logwin.print(f"Success! No-Red CV: {cleaned_cv}", text_color="green")

# ---------------- Tab 3: Three Files (+ Splitter) ----------------
def tab3_process(values, logwin):
    cv = Path((values.get("-T3-CV-") or "").strip())
    master_docx = Path((values.get("-T3-MASTER-") or "").strip())
    sorted_docx = Path((values.get("-T3-SORTED-") or "").strip())
    do_split = bool(values.get("-T3-SPLIT-"))

    if not cv.is_file() or cv.suffix.lower() != ".docx":
        sg.popup_error("Select a valid ORIGINAL CV (.docx)")
        return
    if not master_docx.is_file() or master_docx.suffix.lower() != ".docx":
        sg.popup_error("Select a valid MASTER red-label .docx")
        return
    if not sorted_docx.is_file() or sorted_docx.suffix.lower() != ".docx":
        sg.popup_error("Select a valid SORTED no-red .docx")
        return

    final_cv = OUT / cv.name
    merged_docx = MERGED_DOCX

    args = [exe(), norm(SCRIPT_MERGE),
            "--existing-docx", norm(sorted_docx),
            "--master-docx",   norm(master_docx),
            "--out-docx",      norm(merged_docx),
            "--indent",        "0.5"]
    if not run_cmd(logwin, args):
        return

    args = [exe(), norm(SCRIPT_INJECT),
            "--original-cv",   norm(cv),
            "--studies-docx",  norm(merged_docx),
            "--out",           norm(final_cv),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(logwin, args):
        return

    if do_split:
        args = [exe(), norm(SCRIPT_SPLIT), "--outdir", norm(OUT), norm(final_cv)]
        if not run_cmd(logwin, args):
            return
        move_split_outputs_to_out(final_cv, OUT, logwin)

    for p in [merged_docx]:
        try:
            if p.exists():
                p.unlink()
        except Exception as e:
            logwin.print(f"Could not delete {p}: {e}", text_color="red")

    logwin.print(f"Success! Final CV: {final_cv}", text_color="green")

# ---------------- Main Window ----------------
def main():
    sg.theme("DarkBlue3")

    layout = [
        [sg.TabGroup(
            [[
                sg.Tab("1) All-in-One", tab1_layout(), key="-TAB1-"),
                sg.Tab("2) Remove Red", tab2_layout(), key="-TAB2-"),
                sg.Tab("3) Three Files", tab3_layout(), key="-TAB3-"),
            ]],
            expand_x=True, expand_y=True
        )]
    ]

    window = sg.Window(
        "CenExel CV Pipeline — All-in-One",
        layout,
        resizable=True,
        finalize=True,
        icon=None,
    )

    while True:
        ev, val = window.read()
        if ev in (sg.WIN_CLOSED, "Exit"):
            break

        if ev == "-T1-OPEN-":
            os.startfile(str(OUT))
        if ev == "-T2-OPEN-":
            os.startfile(str(OUT))
        if ev == "-T3-OPEN-":
            os.startfile(str(OUT))

        if ev == "-T1-RUN-":
            logwin = window["-T1-LOG-"]
            tab1_process(val, logwin)
        if ev == "-T2-RUN-":
            logwin = window["-T2-LOG-"]
            tab2_process(val, logwin)
        if ev == "-T3-RUN-":
            logwin = window["-T3-LOG-"]
            tab3_process(val, logwin)

    window.close()

if __name__ == "__main__":
    main()
