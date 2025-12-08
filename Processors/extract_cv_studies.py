#!/usr/bin/env python3
"""
extract_cv_studies.py

Purpose
-------
Extract "unsorted studies" from a full Curriculum Vitae .docx by scanning the
"Research Experience" section and stopping at the disclaimer text:

  "By signing this form, I confirm that the information provided is accurate and reflects my current qualifications."

The output is a plain-text file compatible with the existing sorter (sorterv2.py),
i.e., lines that begin with a 4-digit year followed by any continuation lines.

How it works
------------
- Opens a .docx and finds the section whose heading contains "Research Experience" (case-insensitive).
- Collects all paragraphs until it encounters the exact disclaimer text (case-insensitive match).
- Within that range, builds study blocks as:
    * a new block starts when a paragraph starts with a 4-digit year (e.g., 2023)
    * subsequent non-empty paragraphs are appended to the current block as continuations
    * blank lines close the current block (like the original unsorted .txt format)
- Writes one study per block to the output text file, with blank lines separating studies.

CLI
---
python extract_cv_studies.py --cv "Full_CV.docx" --out ".ADD_CV_STUDIES_FROM_DOCX.txt"

Options
-------
--cv              Path to the .docx CV
--out             Where to write the extracted "unsorted" studies .txt
--section-start   (optional) Text to detect the start of the section (default: "Research Experience")
--section-end     (optional) Text to detect the end of the section (default: the disclaimer sentence above)
--keep-empty      (flag) Keep empty lines inside blocks (default: False)

Notes
-----
- Requires python-docx.
- This script ONLY produces the "unsorted" studies .txt. Feed that into your existing sorterv2.py.
"""

import argparse
import os
import re
from typing import List, Optional

YEAR_RE = re.compile(r'^\s*(\d{4})\b')
DEFAULT_SECTION_START = "Research Experience"
DEFAULT_SECTION_END = "By signing this form, I confirm that the information provided is accurate and reflects my current qualifications."

def _norm(s: str) -> str:
    return ' '.join((s or '').strip().lower().split())

def _para_texts(doc):
    # Yield paragraph texts, preserving order
    for p in doc.paragraphs:
        yield p.text or ''

def _find_bounds(paragraphs: List[str], start_marker: str, end_marker: str) -> Optional[tuple]:
    start_norm = _norm(start_marker)
    end_norm = _norm(end_marker)

    start_idx = None
    for i, t in enumerate(paragraphs):
        if start_norm in _norm(t):
            start_idx = i
            break
    if start_idx is None:
        return None

    end_idx = None
    for j in range(start_idx + 1, len(paragraphs)):
        if end_norm == _norm(paragraphs[j]):
            end_idx = j
            break
    if end_idx is None:
        end_idx = len(paragraphs)
    return (start_idx, end_idx)

def extract_unsorted_from_cv(cv_path: str,
                             out_txt: str,
                             section_start: str = DEFAULT_SECTION_START,
                             section_end: str = DEFAULT_SECTION_END,
                             keep_empty: bool = False) -> None:
    try:
        from docx import Document
    except Exception as e:
        raise RuntimeError("python-docx is required to parse .docx files") from e

    if not os.path.isfile(cv_path):
        raise FileNotFoundError(f"CV .docx not found: {cv_path}")

    doc = Document(cv_path)
    paras = list(_para_texts(doc))
    bounds = _find_bounds(paras, section_start, section_end)
    if bounds is None:
        raise RuntimeError(f'Could not locate section start containing "{section_start}" in the CV.')

    s, e = bounds
    region = paras[s+1:e]  # skip the "Research Experience" heading itself

    studies: List[str] = []
    current: Optional[str] = None

    def flush():
        nonlocal current
        if current is not None and current.strip():
            studies.append(' '.join(current.split()))
        current = None

    for raw in region:
        line = (raw or '').rstrip()
        if not line.strip():
            # blank line: close current block
            flush()
            continue

        if YEAR_RE.match(line):
            # new study starts
            flush()
            current = line.strip()
        else:
            # continuation line (only if a study has started)
            if current is not None:
                if keep_empty:
                    current = f"{current}\n{line.strip()}"
                else:
                    current = f"{current} {line.strip()}"
            else:
                # ignore lines before the first year
                continue

    flush()

    with open(out_txt, 'w', encoding='utf-8') as f:
        for sline in studies:
            f.write(sline + "\n\n")

def main():
    ap = argparse.ArgumentParser(description="Extract unsorted studies from a full CV .docx (Research Experience section).")
    ap.add_argument('--cv', required=True, help='Path to the full CV .docx')
    ap.add_argument('--out', required=True, help='Output path for unsorted studies .txt')
    ap.add_argument('--section-start', default=DEFAULT_SECTION_START, help='Section start marker (default: "Research Experience")')
    ap.add_argument('--section-end', default=DEFAULT_SECTION_END, help='Section end marker (default: the standard disclaimer sentence)')
    ap.add_argument('--keep-empty', action='store_true', help='Keep internal empty lines inside a study block')
    args = ap.parse_args()

    extract_unsorted_from_cv(args.cv, args.out, args.section_start, args.section_end, args.keep_empty)
    print("Done.")
    print(f"  Extracted unsorted studies to: {os.path.abspath(args.out)}")

if __name__ == '__main__':
    main()
