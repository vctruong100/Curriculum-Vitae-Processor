#!/usr/bin/env python3
# GUI for CSV-mapped fuzzy replacement (no color checks)
# Output filename is forced to "<original name> (UPDATED).docx"

import os, sys, subprocess
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

SCRIPT_REMOVE = HERE / "remove_red_labels_from_docx.py"

SECTION_START  = "Research Experience"
SECTION_END    = "By signing this form, I confirm that the information provided is accurate and reflects my current qualifications."

def exe():
    return sys.executable or "python"

def norm(p: Path) -> str:
    return str(p.expanduser().resolve())

def run_cmd(window, args):
    window["-LOG-"].print(f"$ {' '.join(args)}", text_color="yellow")
    try:
        p = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        for line in p.stdout:
            window["-LOG-"].print(line.rstrip())
        rc = p.wait()
        if rc != 0:
            window["-LOG-"].print(f"Exited with code {rc}", text_color="red")
            sg.popup_error("Step failed. See log for details.")
            return False
        window["-LOG-"].print("Done.", text_color="green")
        return True
    except Exception as e:
        window["-LOG-"].print(f"ERROR: {e}", text_color="red")
        sg.popup_error(f"ERROR: {e}")
        return False

def make_updated_name(src: Path) -> Path:
    # Keep original name, append " (UPDATED).docx"
    stem = src.stem
    return OUT / f"{stem} (UPDATED){src.suffix}"

sg.theme("SystemDefault")
layout = [
    [sg.Text("Remove Red Labels — CSV Fuzzy Matching (no re-sorting)", font=("Segoe UI", 11))],
    [sg.Text("Final CV (.docx)", size=(22,1)), sg.Input(key="-CV-", expand_x=True), sg.FileBrowse(file_types=(("Word", "*.docx"),))],
    [sg.Text("Mapping CSV", size=(22,1)), sg.Input(key="-CSV-", expand_x=True), sg.FileBrowse(file_types=(("CSV", "*.csv"),))],
    [sg.Text("Fuzzy threshold (0.80–0.98)"), sg.Slider(range=(0.80,0.98), resolution=0.01, default_value=0.90, orientation="h", size=(30,20), key="-TH-")],
    [sg.Button("Process", bind_return_key=True, size=(16,1)), sg.Button("Open Output"), sg.Button("Quit")],
    [sg.Frame("Log", [[sg.Multiline(size=(100,18), key="-LOG-", autoscroll=True, expand_x=True, expand_y=True, write_only=True)]], expand_x=True)],
]
window = sg.Window("CV Pipeline — Remove Red Labels (CSV Fuzzy)", layout, finalize=True)

while True:
    ev, val = window.read()
    if ev in (sg.WIN_CLOSED, "Quit"):
        break
    if ev == "Open Output":
        os.startfile(str(OUT))
    if ev == "Process":
        cv = Path((val.get("-CV-") or "").strip())
        csvp = Path((val.get("-CSV-") or "").strip())
        if not cv.is_file() or cv.suffix.lower() != ".docx":
            sg.popup_error("Select a valid Final CV (.docx)")
            continue
        if not csvp.is_file() or csvp.suffix.lower() != ".csv":
            sg.popup_error("Select a valid mapping CSV (.csv)")
            continue
        out_cv = make_updated_name(cv)
        args = [exe(), norm(SCRIPT_REMOVE),
                "--original-cv", str(cv),
                "--mapping-csv", str(csvp),
                "--out",         str(out_cv),
                "--section-start", SECTION_START,
                "--section-end",   SECTION_END,
                "--threshold",     f"{float(val['-TH-']):.2f}"]
        run_cmd(window, args)

window.close()
