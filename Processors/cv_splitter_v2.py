
#!/usr/bin/env python3
"""
cv_splitter_v2.py — Split AFTER the **first** signature block, then export PDFs.

- Keeps original behavior: split on the FIRST signature/Date-of-Signature block.
- Outputs go to --outdir (authoritative), else OUTPUT_DIR env, else CWD.
- Robust docx→pdf with per-file handling and relocation into outdir if Word drops elsewhere.
"""

import os, sys, argparse, shutil
from typing import List, Optional, Tuple
from docx import Document

try:
    from docx2pdf import convert
except Exception:
    convert = None  # allow running without Word; will keep DOCX

SIGN_LABEL_SUBSTR = "by signing this form"
SIG_CELL_1 = "signature"
SIG_CELL_2 = "date of signature"


def ensure_dir(path: str):
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass


def ensure_parent(path: str):
    ensure_dir(os.path.dirname(os.path.abspath(path)) or os.getcwd())


def body_elements(doc: Document):
    body = doc.element.body
    for child in list(body.iterchildren()):
        tag = child.tag.rsplit('}', 1)[-1]
        if tag == 'p':
            yield 'p', child
        elif tag == 'tbl':
            yield 'tbl', child
        else:
            yield tag, child


def paragraph_text_from_xml(p_xml) -> str:
    texts: List[str] = []
    for r in p_xml.iterchildren():
        if r.tag.rsplit('}', 1)[-1] == 'r':
            for t in r.iterchildren():
                if t.tag.rsplit('}', 1)[-1] == 't' and t.text:
                    texts.append(t.text)
    return ''.join(texts)


def cell_text_from_tc(tc_xml) -> str:
    parts: List[str] = []
    for p in tc_xml.iterchildren():
        if p.tag.rsplit('}', 1)[-1] == 'p':
            t = paragraph_text_from_xml(p)
            if t:
                parts.append(t)
    txt = ' '.join(parts).strip().lower()
    txt = txt.replace('_', ' ')
    txt = ' '.join(txt.split())
    return txt


def table_contains_signature_markers(tbl_xml) -> bool:
    sig_found = False
    date_found = False
    for row in tbl_xml.iterchildren():
        if row.tag.rsplit('}', 1)[-1] != 'tr':
            continue
        for cell in row.iterchildren():
            if cell.tag.rsplit('}', 1)[-1] != 'tc':
                continue
            txt = cell_text_from_tc(cell)
            if SIG_CELL_1 in txt:
                sig_found = True
            if SIG_CELL_2 in txt:
                date_found = True
            if sig_found and date_found:
                return True
    return False


def find_first_signature_table_index(doc: Document) -> Optional[int]:
    elems = list(body_elements(doc))
    # Pass 1: paragraph "By signing..." then the next table w/ markers
    for i, (tag, xml) in enumerate(elems):
        if tag != 'p':
            continue
        text = paragraph_text_from_xml(xml).strip().lower()
        if SIGN_LABEL_SUBSTR in text:
            for j in range(i + 1, len(elems)):
                t, x = elems[j]
                if t == 'tbl' and table_contains_signature_markers(x):
                    return j
            break
    # Pass 2: first signature table anywhere
    for i, (tag, xml) in enumerate(elems):
        if tag == 'tbl' and table_contains_signature_markers(xml):
            return i
    return None


def remove_body_range(doc: Document, start_idx: int, end_idx: int) -> None:
    if start_idx < 0:
        start_idx = 0
    children = list(doc.element.body.iterchildren())
    end_idx = min(end_idx, len(children) - 1)
    if start_idx > end_idx:
        return
    for idx in range(start_idx, end_idx + 1):
        child = children[idx]
        parent = child.getparent()
        if parent is not None:
            parent.remove(child)


def remove_prefix(doc: Document, count: int) -> None:
    if count <= 0:
        return
    children = list(doc.element.body.iterchildren())
    end_idx = min(count - 1, len(children) - 1)
    remove_body_range(doc, 0, end_idx)


def remove_suffix_after(doc: Document, idx: int) -> None:
    children = list(doc.element.body.iterchildren())
    last = len(children) - 1
    if idx < last:
        remove_body_range(doc, idx + 1, last)


def disable_different_first_page(doc: Document) -> None:
    # Ensure header/footer show on page 1 for all sections
    for sec in doc.sections:
        try:
            sec.different_first_page_header_footer = False
        except Exception:
            pass


def has_only_pagebreaks_or_whitespace(p_xml) -> bool:
    txt = paragraph_text_from_xml(p_xml)
    if txt and txt.strip():
        return False
    for r in p_xml.iterchildren():
        if r.tag.rsplit('}', 1)[-1] != 'r':
            continue
        for el in r.iterchildren():
            if el.tag.rsplit('}', 1)[-1] == 'br':
                return True
    return True


def strip_leading_blank_or_pagebreak_paragraphs(doc: Document) -> None:
    while True:
        children = list(doc.element.body.iterchildren())
        if not children:
            break
        first = children[0]
        tag = first.tag.rsplit('}', 1)[-1]
        if tag != 'p':
            break
        if has_only_pagebreaks_or_whitespace(first):
            parent = first.getparent()
            if parent is not None:
                parent.remove(first)
            continue
        break


def parse_person_from_filename(path: str) -> str:
    # Use everything after "Template " if present, else the tail (no extension)
    base = os.path.basename(path)
    if base.lower().endswith('.docx'):
        base = base[:-5]
    idx = base.lower().rfind('template ')
    if idx != -1:
        name = base[idx + len('template '):].strip()
        return name or base
    return base


def make_out_paths(person: str, outdir: str) -> Tuple[str, str, str, str]:
    out_full_docx = os.path.join(outdir, f"CenExel CURRICULUM VITAE {person}.docx")
    out_abbr_docx = os.path.join(outdir, f"CenExel Abbrv CURRICULUM VITAE {person}.docx")
    out_full_pdf = os.path.splitext(out_full_docx)[0] + ".pdf"
    out_abbr_pdf = os.path.splitext(out_abbr_docx)[0] + ".pdf"
    for p in (out_full_docx, out_abbr_docx, out_full_pdf, out_abbr_pdf):
        ensure_parent(p)
    return out_full_docx, out_abbr_docx, out_full_pdf, out_abbr_pdf


def _ensure_in_outdir(pdf_path: str, outdir: str) -> str:
    try:
        if os.path.exists(pdf_path):
            abs_pdf = os.path.abspath(pdf_path)
            if os.path.dirname(abs_pdf).lower() == os.path.abspath(outdir).lower():
                return abs_pdf
            dest = os.path.join(outdir, os.path.basename(abs_pdf))
            try:
                if os.path.exists(dest):
                    os.remove(dest)
            except Exception:
                pass
            shutil.move(abs_pdf, dest)
            return os.path.abspath(dest)
    except Exception:
        pass
    return os.path.abspath(pdf_path)


def _convert_one(src_docx: str, dst_pdf: str, label: str, outdir: str) -> bool:
    if convert is None:
        print(f"[INFO] docx2pdf not available. Kept DOCX for {label}: {os.path.abspath(src_docx)}")
        return False
    try:
        convert(src_docx, dst_pdf)
    except Exception as e:
        if os.path.exists(dst_pdf):
            print(f"[WARN] {label}: Exception during convert() but PDF exists -> treating as success. ({e})")
        else:
            print(f"[WARN] {label}: PDF conversion failed: {e}")
            return False
    if os.path.exists(dst_pdf):
        final_pdf = _ensure_in_outdir(dst_pdf, outdir)
        try:
            os.remove(src_docx)
        except Exception:
            pass
        print(f"  {label} (PDF): {final_pdf}")
        return True
    else:
        print(f"[WARN] {label}: No PDF produced; kept DOCX: {os.path.abspath(src_docx)}")
        return False


def main():
    ap = argparse.ArgumentParser(description="Split CV around FIRST signature block and convert to PDF.")
    ap.add_argument("--outdir", default=None, help="Directory to write outputs")
    ap.add_argument("input_cv", help="Path to the final updated CV .docx")
    args = ap.parse_args()

    outdir = args.outdir or os.environ.get("OUTPUT_DIR") or os.getcwd()
    outdir = os.path.abspath(outdir)
    ensure_dir(outdir)

    in_path = os.path.abspath(args.input_cv)
    if not os.path.isfile(in_path):
        print(f"ERROR: File not found: {in_path}")
        sys.exit(2)

    print(f"Input CV: {in_path}")
    person = parse_person_from_filename(in_path)
    out_full_docx, out_abbr_docx, out_full_pdf, out_abbr_pdf = make_out_paths(person, outdir)

    # Load original twice so each output keeps relationships intact
    doc_abbr = Document(in_path)
    doc_full = Document(in_path)

    idx = find_first_signature_table_index(doc_abbr)
    idx2 = find_first_signature_table_index(doc_full)
    if idx is None or idx2 is None:
        print("ERROR: Could not locate the Signature / Date of Signature box.")
        sys.exit(1)

    # Abbreviated: keep through the signature table
    remove_suffix_after(doc_abbr, idx)

    # Full: keep content AFTER the signature table
    remove_prefix(doc_full, idx2 + 1)

    # Fix: headers/footers visible on first page, remove leading blanks
    disable_different_first_page(doc_abbr)
    disable_different_first_page(doc_full)
    strip_leading_blank_or_pagebreak_paragraphs(doc_full)

    # Save DOCX
    doc_abbr.save(out_abbr_docx)
    doc_full.save(out_full_docx)

    # Convert DOCX -> PDF (per-file robustness)
    ok_abbr = _convert_one(out_abbr_docx, out_abbr_pdf, "Abbreviated CV", outdir)
    ok_full = _convert_one(out_full_docx, out_full_pdf, "Full CV", outdir)

    if ok_abbr and ok_full:
        print("Converted to PDF successfully.")
        sys.exit(0)
    elif ok_abbr or ok_full:
        print("[PARTIAL] One of the PDFs was created; see messages above.")
        sys.exit(0)
    else:
        print("[FAIL] No PDFs created; kept DOCX outputs (see messages above).")
        print(f"  Abbreviated CV (DOCX): {os.path.abspath(out_abbr_docx)}")
        print(f"  Full CV        (DOCX): {os.path.abspath(out_full_docx)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
