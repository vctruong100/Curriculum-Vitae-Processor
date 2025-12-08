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

SECTION_START  = "Research Experience"
SECTION_END    = "By signing this form, I confirm that the information provided is accurate and reflects my current qualifications."

UNSORTED_TXT   = OUT / ".ADD_CV_STUDIES_FROM_DOCX.txt"
SORTED_TXT     = OUT / "SORTED_STUDY_CV_TXT.txt"
AUDIT_TSV      = OUT / "match_audit_report.tsv"
NOYEAR_AUDIT   = OUT / "noyear_resolve_audit.tsv"
SORTED_DOCX    = OUT / "SORTED_STUDY_CV_DOCX.docx"
MERGED_DOCX    = OUT / ".UPDATED CV.docx"

def exe():
    return sys.executable or "python"

def norm(p: Path) -> str:
    return str(p.expanduser().resolve())

def run_cmd(window, args, cwd=None):
    window.print(f"$ {' '.join(args)}", text_color="yellow")
    try:
        p = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, cwd=cwd)
        for line in p.stdout:
            window.print(line.rstrip())
        rc = p.wait()
        if rc != 0:
            window.print(f"Exited with code {rc}", text_color="red")
            sg.popup_error("A step failed. See log for details.")
            return False
        window.print("Done.", text_color="green")
        return True
    except Exception as e:
        window.print(f"ERROR: {e}", text_color="red")
        sg.popup_error(f"ERROR: {e}")
        return False

def move_split_outputs_to_out(final_cv_path: Path, out_dir: Path, window):
    patterns = [
        "CenExel CURRICULUM VITAE",
        "CenExel Abbrv CURRICULUM VITAE"
    ]
    roots_to_scan = [final_cv_path.parent, ROOT, OUT]
    candidates = []
    for root_dir in roots_to_scan:
        try:
            for f in list(root_dir.glob("*.docx")) + list(root_dir.glob("*.pdf")):
                name = f.name
                if any(name.startswith(pat) for pat in patterns):
                    candidates.append(f)
        except Exception:
            pass
    moved = []
    for f in candidates:
        dest = out_dir / f.name
        try:
            if f.resolve() != dest.resolve():
                dest.write_bytes(f.read_bytes())
                f.unlink()
                moved.append(str(dest))
        except Exception as e:
            window.print(f"Could not move {f} -> {dest}: {e}", text_color="red")
    if moved:
        window.print("Moved split outputs to Output/:")
        for m in moved:
            window.print(f"  {m}")

# ---------------- Tab 1: One-Click (+ Splitter) with CSV No-Year Fix ----------------
def tab1_layout():
    return [
        [sg.Text("Original CV (.docx)", size=(22,1)), sg.Input(key="-T1-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Text("Mapping CSV (A=Year, C=Non-Red study)", size=(22,1)), sg.Input(key="-T1-CSV-", expand_x=True), sg.FileBrowse(file_types=(("CSV", "*.csv"),))],
        [sg.Text("Fuzzy match threshold"), sg.Input(key="-T1-THRESH-", size=(8,1), default_text="0.88")],
        [sg.Checkbox("Also create Abbreviated & Full CV from the FINAL CV (run splitter)", key="-T1-SPLIT-", default=False)],
        [sg.Button("Run One-Click", key="-T1-RUN-", size=(18,1)), sg.Button("Open Output", key="-T1-OPEN-")],
        [sg.Frame("Log", [[sg.Multiline(size=(100,18), key="-T1-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)],
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

    # 1) Extract
    args = [exe(), norm(SCRIPT_EXTRACT),
            "--cv", norm(cv),
            "--out", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(logwin, args): return

    # 2) Resolve no-year using CSV
    args = [exe(), norm(SCRIPT_RESOLVE),
            "--cv", norm(cv),
            "--csv", norm(csvp),
            "--in-unsorted", norm(UNSORTED_TXT),
            "--out-unsorted", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END,
            "--threshold",     str(th),
            "--audit",         norm(NOYEAR_AUDIT)]
    if not run_cmd(logwin, args): return

    # 3) Sort
    args = [exe(), norm(SCRIPT_SORT),
            "--master",        str(Path("Editable") / ".NO_RED_STUDYLIST (EDITABLE).txt"),
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
    if not run_cmd(logwin, args): return

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
            if not run_cmd(logwin, args): return
            use_for_inject = MERGED_DOCX
            break

    # 5) Inject
    args = [exe(), norm(SCRIPT_INJECT),
            "--original-cv",   norm(cv),
            "--studies-docx",  norm(use_for_inject),
            "--out",           norm(final_cv),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(logwin, args): return

    # 6) Optional splitter
    if do_split:
        args = [exe(), norm(SCRIPT_SPLIT), "--outdir", norm(OUT), norm(final_cv)]
        if not run_cmd(logwin, args): return
        move_split_outputs_to_out(final_cv, OUT, logwin)

    # Cleanup
    for p in [UNSORTED_TXT, SORTED_TXT, AUDIT_TSV, NOYEAR_AUDIT, SORTED_DOCX, MERGED_DOCX]:
        try:
            if p.exists():
                p.unlink()
        except Exception as e:
            logwin.print(f"Could not delete {p}: {e}", text_color="red")

    logwin.print(f"Success! Final CV: {final_cv}", text_color="green")

# ---------------- Tab 2: Remove Red Labels (CSV Fuzzy) + Splitter ----------------
def tab2_layout():
    return [
        [sg.Text("Final CV (.docx)", size=(22,1)), sg.Input(key="-T2-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Text("Mapping CSV", size=(22,1)), sg.Input(key="-T2-CSV-", expand_x=True), sg.FileBrowse(file_types=(("CSV", "*.csv"),))],
        [sg.Text("Fuzzy threshold (0.80–0.98)"), sg.Slider(range=(0.80,0.98), resolution=0.01, default_value=0.90, orientation="h", size=(30,20), key="-T2-TH-")],
        [sg.Checkbox("Also create Abbreviated & Full CV from the UPDATED CV (run splitter)", key="-T2-SPLIT-", default=True)],
        [sg.Button("Run Remove-Red", key="-T2-RUN-", size=(18,1)), sg.Button("Open Output", key="-T2-OPEN-")],
        [sg.Frame("Log", [[sg.Multiline(size=(100,18), key="-T2-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)],
    ]

def tab2_process(values, logwin):
    cv = Path((values.get("-T2-CV-") or "").strip())
    csvp = Path((values.get("-T2-CSV-") or "").strip())
    th = float(values.get("-T2-TH-", 0.90))
    do_split = bool(values.get("-T2-SPLIT-", True))

    if not cv.is_file() or cv.suffix.lower() != ".docx":
        sg.popup_error("Select a valid Final CV (.docx)")
        return
    if not csvp.is_file() or csvp.suffix.lower() != ".csv":
        sg.popup_error("Select a valid mapping CSV (.csv)")
        return

    out_cv = OUT / f"{cv.stem} (UPDATED){cv.suffix}"

    # Remove red labels from the provided Final CV using CSV map
    args = [exe(), norm(SCRIPT_REMOVE),
            "--original-cv", str(cv),
            "--mapping-csv", str(csvp),
            "--out",         str(out_cv),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END,
            "--threshold",     f"{th:.2f}"]
    if not run_cmd(logwin, args): return

    # Split into Abbrv/Full PDFs (like other tabs)
    if do_split:
        args = [exe(), norm(SCRIPT_SPLIT), "--outdir", norm(OUT), norm(out_cv)]
        if not run_cmd(logwin, args): return
        move_split_outputs_to_out(out_cv, OUT, logwin)

    logwin.print(f"✓ Done. Output folder: {OUT}", text_color="green")

# ---------------- Tab 3: Three Files (+ Splitter) ----------------
def tab3_layout():
    return [
        [sg.Text("Original CV (.docx)", size=(22,1)), sg.Input(key="-T3-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Text("Unred MASTER (.txt)", size=(22,1)), sg.Input(key="-T3-MTXT-", expand_x=True), sg.FileBrowse(file_types=(("Text", "*.txt"),))],
        [sg.Text("MASTER w/ red (.docx)", size=(22,1)), sg.Input(key="-T3-MRED-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
        [sg.Checkbox("Also create Abbreviated & Full CV from the FINAL CV (run splitter)", key="-T3-SPLIT-", default=True)],
        [sg.Button("Run Three-File", key="-T3-RUN-", size=(18,1)), sg.Button("Open Output", key="-T3-OPEN-")],
        [sg.Frame("Log", [[sg.Multiline(size=(100,18), key="-T3-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)],
    ]

def tab3_process(values, logwin):
    cv_path   = Path((values.get("-T3-CV-") or "").strip())
    master_txt = Path((values.get("-T3-MTXT-") or "").strip())
    master_red = Path((values.get("-T3-MRED-") or "").strip())
    do_split   = bool(values.get("-T3-SPLIT-"))

    if not cv_path.is_file() or cv_path.suffix.lower() != ".docx":
        sg.popup_error("Select a valid ORIGINAL CV .docx")
        return
    if not master_txt.is_file() or master_txt.suffix.lower() != ".txt":
        sg.popup_error("Select a valid UNRED MASTER schedule .txt")
        return
    if not master_red.is_file() or master_red.suffix.lower() != ".docx":
        sg.popup_error("Select a valid MASTER schedule with red labels .docx")
        return

    final_cv = OUT / cv_path.name

    # Extract
    args = [exe(), norm(SCRIPT_EXTRACT),
            "--cv", norm(cv_path),
            "--out", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(logwin, args): return

    # Sort
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
    if not run_cmd(logwin, args): return

    # Merge
    args = [exe(), norm(SCRIPT_MERGE),
            "--existing-docx", norm(SORTED_DOCX),
            "--master-docx",   norm(master_red),
            "--out-docx",      norm(MERGED_DOCX),
            "--indent",        "0.5"]
    if not run_cmd(logwin, args): return

    # Inject
    args = [exe(), norm(SCRIPT_INJECT),
            "--original-cv",  norm(cv_path),
            "--studies-docx", norm(MERGED_DOCX),
            "--out",          norm(final_cv),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(logwin, args): return

    if do_split:
        args = [exe(), norm(SCRIPT_SPLIT), "--outdir", norm(OUT), norm(final_cv)]
        if not run_cmd(logwin, args): return
        move_split_outputs_to_out(final_cv, OUT, logwin)

    # Cleanup
    for p in [UNSORTED_TXT, SORTED_TXT, AUDIT_TSV, SORTED_DOCX, MERGED_DOCX]:
        try:
            if p.exists():
                p.unlink()
        except Exception as e:
            logwin.print(f"Could not delete {p}: {e}", text_color="red")

    logwin.print(f"✓ Done. Output folder: {OUT}", text_color="green")

# -------------- Build the Tabbed UI --------------
sg.theme("SystemDefault")

tab1 = sg.Tab("One-Click (CSV No-Year Fix)", tab1_layout(), key="-TAB1-")
tab2 = sg.Tab("Remove Red Labels (CSV)",     tab2_layout(), key="-TAB2-")
tab3 = sg.Tab("Three Files (+ Splitter)",    tab3_layout(), key="-TAB3-")

layout = [
    [sg.TabGroup([[tab1, tab2, tab3]], key="-TABS-", expand_x=True, expand_y=True)],
    [sg.Button("Quit")]
]

window = sg.Window("CV Pipeline — All-in-One", layout, finalize=True, resizable=True)

while True:
    ev, val = window.read()
    if ev in (sg.WIN_CLOSED, "Quit"):
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
