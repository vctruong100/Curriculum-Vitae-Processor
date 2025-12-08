
#!/usr/bin/env python3
# CV Pipeline — One Click (+ Splitter) with CSV fix for no-year studies
# Adds a CSV step between Extract and Sort. Keeps all original filenames unchanged.
#
# Outputs remain:
#   UNSORTED_TXT   = Output/.ADD_CV_STUDIES_FROM_DOCX.txt (updated in-place with resolved lines)
#   SORTED_TXT     = Output/SORTED_STUDY_CV_TXT.txt
#   AUDIT_TSV      = Output/match_audit_report.tsv  (from sorter)
#   SORTED_DOCX    = Output/SORTED_STUDY_CV_DOCX.docx
#   MERGED_DOCX    = Output/.UPDATED CV.docx
#   Final injected = Output/<original name>.docx
#   Split (optional): CenExel CURRICULUM VITAE {Name}.docx, CenExel Abbrv CURRICULUM VITAE {Name}.docx

import os, sys, subprocess
import argparse
from pathlib import Path

try:
    import FreeSimpleGUI as sg
except Exception:
    try:
        import PySimpleGUI as sg
    except Exception:
        raise SystemExit("FreeSimpleGUI not installed. Run via the .bat launcher to auto-install.")

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
OUT  = ROOT / "Output"
OUT.mkdir(parents=True, exist_ok=True)

# Processor scripts (must live in the same folder as this GUI)
SCRIPT_EXTRACT = HERE / "extract_cv_studies.py"
SCRIPT_RESOLVE = HERE / "resolve_noyear_from_csv.py"
SCRIPT_SORT    = HERE / "sorterv2.py"
SCRIPT_MERGE   = HERE / "compare_insert_red_docx.py"
SCRIPT_INJECT  = HERE / "inject_sorted_into_cv.py"
SCRIPT_SPLIT   = HERE / "cv_splitter_v2.py"

SECTION_START  = "Research Experience"
SECTION_END    = "By signing this form, I confirm that the information provided is accurate and reflects my current qualifications."

# Intermediates / outputs
UNSORTED_TXT   = OUT / ".ADD_CV_STUDIES_FROM_DOCX.txt"
SORTED_TXT     = OUT / "SORTED_STUDY_CV_TXT.txt"
AUDIT_TSV      = OUT / "match_audit_report.tsv"
NOYEAR_AUDIT   = OUT / "noyear_resolve_audit.tsv"
SORTED_DOCX    = OUT / "SORTED_STUDY_CV_DOCX.docx"
MERGED_DOCX    = OUT / ".UPDATED CV.docx"

# CLI arg: allow overriding threshold (default 0.88)
DEFAULT_THRESHOLD = 0.88
ap = argparse.ArgumentParser(add_help=False)
ap.add_argument('--threshold', type=float, default=DEFAULT_THRESHOLD)
cli_args, _ = ap.parse_known_args()
CLI_THRESHOLD = cli_args.threshold if (cli_args.threshold is not None) else DEFAULT_THRESHOLD

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


def coerce_threshold(val, fallback: float) -> float:
    try:
        t = float(str(val).strip())
        if 0.0 <= t <= 1.0:
            return t
    except Exception:
        pass
    return fallback


def move_split_outputs_to_out(final_cv_path: Path, out_dir: Path, window):
    """
    Ensure splitter-produced files are saved under Output/.
    We look for likely names in both the CV's directory and the project root.
    """
    # Expected output names contain the person's name; we search for patterns.
    candidates = []
    roots_to_scan = [final_cv_path.parent, ROOT, OUT]
    patterns = [
        "CenExel CURRICULUM VITAE",
        "CenExel Abbrv CURRICULUM VITAE"
    ]
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
            window["-LOG-"].print(f"Could not move {f} -> {dest}: {e}", text_color="red")
    if moved:
        window["-LOG-"].print("Moved split outputs to Output/:")
        for m in moved:
            window["-LOG-"].print(f"  {m}")


sg.theme("SystemDefault")
layout = [
    [sg.Text("One-Click CV Pipeline (+ Splitter) — with CSV No-Year Fix", font=("Segoe UI", 11))],
    [sg.Text("Original CV (.docx)", size=(22,1)), sg.Input(key="-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
    [sg.Text("Mapping CSV (A=Year, C=Non-Red study)", size=(22,1)), sg.Input(key="-CSV-", expand_x=True), sg.FileBrowse(file_types=(("CSV", "*.csv"),))],
    [sg.Text("Fuzzy match threshold", size=(22,1)), sg.Input(key="-THRESH-", size=(8,1), default_text=str(CLI_THRESHOLD))],
    [sg.Checkbox("Also create Abbreviated & Full CV from the FINAL CV (run splitter)", key="-DO_SPLIT-", default=False)],
    [sg.Button("Process", bind_return_key=True, size=(18,1)), sg.Button("Open Output"), sg.Button("Quit")],
    [sg.Frame("Log", [[sg.Multiline(size=(100,22), key="-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)]
]
window = sg.Window("CV Pipeline — One Click (No-Year CSV Fix)", layout, finalize=True)

def process_all(window, cv_file: str, csv_file: str, do_split: bool, threshold: float):
    cv_path  = Path((cv_file or '').strip())
    csv_path = Path((csv_file or '').strip())
    if not cv_path.is_file() or cv_path.suffix.lower() != ".docx":
        sg.popup_error("Select a valid ORIGINAL CV (.docx)")
        return
    if not csv_path.is_file() or csv_path.suffix.lower() != ".csv":
        sg.popup_error("Select a valid mapping CSV (.csv)")
        return

    final_cv = OUT / cv_path.name  # keep the same filename as the original

    window["-LOG-"].print(f"Original CV: {cv_path}")
    window["-LOG-"].print(f"Mapping CSV: {csv_path}")

    # 1) Extract unsorted (yeared-only) from the CV into UNSORTED_TXT
    args = [exe(), norm(SCRIPT_EXTRACT),
            "--cv", norm(cv_path),
            "--out", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(window, args): return

    # 2) Resolve no-year lines using CSV, MERGE back into the same UNSORTED_TXT
    args = [exe(), norm(SCRIPT_RESOLVE),
            "--cv", norm(cv_path),
            "--csv", norm(csv_path),
            "--in-unsorted", norm(UNSORTED_TXT),
            "--out-unsorted", norm(UNSORTED_TXT),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END,
            "--threshold",     str(threshold),
            "--audit",         norm(NOYEAR_AUDIT)]
    if not run_cmd(window, args): return

    # 3) Sort using master .txt (auto-detected inside sorterv2 step)
    #    NOTE: sorterv2.py expects --master path to the NO-RED master; your existing GUI(s) locate it automatically.
    #    Here we mimic the same call signature as cv_gui_oneclick_plus_splitter.py.
    #    You will run this GUI from within your existing folder where sorterv2.py already locates the MASTER (same behavior as before).
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
    if not run_cmd(window, args): return

    # 4) Merge to preserve red labels, if a MASTER red .docx exists (auto-detect like your other GUI)
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
            window["-LOG-"].print(f"MASTER .docx found: {c} → merging to preserve red labels")
            args = [exe(), norm(SCRIPT_MERGE),
                    "--existing-docx", norm(SORTED_DOCX),
                    "--master-docx",   norm(c),
                    "--out-docx",      norm(MERGED_DOCX),
                    "--indent",        "0.5"]
            if not run_cmd(window, args): return
            use_for_inject = MERGED_DOCX
            break

    # 5) Inject merged into cloned original (final file keeps the SAME name in Output/)
    args = [exe(), norm(SCRIPT_INJECT),
            "--original-cv",   norm(cv_path),
            "--studies-docx",  norm(use_for_inject),
            "--out",           norm(final_cv),
            "--section-start", SECTION_START,
            "--section-end",   SECTION_END]
    if not run_cmd(window, args): return

    # 6) Optional Splitter from FINAL CV
    if do_split:
        args = [exe(), norm(SCRIPT_SPLIT), "--outdir", norm(OUT), norm(final_cv)]
        if not run_cmd(window, args): return
        move_split_outputs_to_out(final_cv, OUT, window)

    window["-LOG-"].print(f"Success! Final CV: {final_cv}", text_color="green")
    # Remove intermediates so only the final(s) remain
    keep = {str(final_cv)}
    # If split is requested, also preserve the two split outputs (ensured under Output/)
    if do_split:
        # Defer: move_split_outputs_to_out already moved files; just track against deletion
        pass
    # Intermediates to delete:
    for pth in [UNSORTED_TXT, SORTED_TXT, AUDIT_TSV, NOYEAR_AUDIT, SORTED_DOCX, MERGED_DOCX]:
        try:
            if pth.exists():
                pth.unlink()
                window["-LOG-"].print(f"Deleted: {pth}")
        except Exception as _e:
            window["-LOG-"].print(f"Could not delete {pth}: {_e}")

    sg.popup("Done!", title="CV Pipeline (No-Year CSV Fix)")

while True:
    ev, val = window.read()
    if ev in (sg.WIN_CLOSED, "Quit"):
        break
    if ev == "Open Output":
        os.startfile(str(OUT))
    if ev == "Process":
        t = coerce_threshold(val.get("-THRESH-"), CLI_THRESHOLD)
        process_all(window, val.get("-CV-"), val.get("-CSV-"), val.get("-DO_SPLIT-"), t)
