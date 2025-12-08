#!/usr/bin/env python3
"""
Split a CV .docx into two outputs around the first Signature/Date signature box,
then export both outputs to PDF. If no input path is provided, the script will
auto-find the newest suitable .docx in the current folder or in
"Editable (Add CV here)".

Input name example:
  "CenExel CURRICULUM VITAE Template Patrick Atiyuthkul.docx"

PDF Outputs:
  1) CenExel CURRICULUM VITAE {Name}.pdf
  2) CenExel Abbrv CURRICULUM VITAE {Name}.pdf

Changes vs v2 base:
  - Auto-detect input .docx when no CLI arg is supplied
  - Force headers/footers to appear on the first page (disable "Different First Page")
  - Remove leading empty/page-break paragraphs on the Full CV to avoid a blank first page
  - Convert final .docx files to .pdf via docx2pdf (and remove the .docx if conversion succeeds)
"""

import os
import sys
from typing import List, Optional, Tuple
from docx import Document

try:
    from docx2pdf import convert
except Exception as e:
    convert = None  # We'll handle missing docx2pdf at runtime


SIGN_LABEL_SUBSTR = "by signing this form"
SIG_CELL_1 = "signature"
SIG_CELL_2 = "date of signature"


def parse_person_name_from_filename(path: str) -> str:
    base = os.path.basename(path)
    if base.lower().endswith('.docx'):
        base = base[:-5]
    idx = base.lower().rfind('template ')
    if idx != -1:
        name = base[idx + len('template '):].strip()
        if name:
            return name

    prefixes = [
        'cenexel curriculum vitae',
        'cenexel abbrv curriculum vitae',
        'cenexel curriculum vitae template'
    ]
    tmp = base
    low = base.lower()
    for p in prefixes:
        if low.startswith(p):
            tmp = base[len(p):].strip(" -_")
            break
    return tmp if tmp else base


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

    # Pass 1: paragraph with "By signing..." then next table w/ markers
    for i, (tag, xml) in enumerate(elems):
        if tag != 'p':
            continue
        text = paragraph_text_from_xml(xml).strip().lower()
        if SIGN_LABEL_SUBSTR in text:
            for j in range(i + 1, len(elems)):
                t, x = elems[j]
                if t == 'tbl':
                    if table_contains_signature_markers(x):
                        return j
            break

    # Pass 2: first table w/ markers anywhere
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
    """
    Return True if paragraph has no visible text and only contains page/line breaks or whitespace.
    """
    txt = paragraph_text_from_xml(p_xml)
    if txt and txt.strip():
        # there is visible text
        return False
    # Look for explicit w:br elements (page/line)
    for r in p_xml.iterchildren():
        if r.tag.rsplit('}', 1)[-1] != 'r':
            continue
        for el in r.iterchildren():
            if el.tag.rsplit('}', 1)[-1] == 'br':
                # a break is present
                return True
    # no text & no breaks -> effectively blank
    return True


def strip_leading_blank_or_pagebreak_paragraphs(doc: Document) -> None:
    # After removing the prefix, the first few paragraphs can be empty or just breaks.
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
            # continue and examine next
            continue
        # first paragraph has visible text -> stop
        break


def is_generated_output_name(filename_lower: str) -> bool:
    # Skip files the script itself would generate
    return (
        filename_lower.startswith("cenexel curriculum vitae ") or
        filename_lower.startswith("cenexel abbrv curriculum vitae ")
    )


def looks_like_template_name(filename_lower: str) -> bool:
    # Prefer files that look like the input template
    return (
        filename_lower.endswith(".docx") and
        "template" in filename_lower and
        "curriculum" in filename_lower and
        "vitae" in filename_lower
    )


def auto_find_input() -> Optional[str]:
    """
    Find the newest plausible input .docx in:
      1) current directory
      2) ./Editable (Add CV here)/
    Preference order:
      - Files that look like "Template ... .docx"
      - Otherwise, any .docx that is not a generated output
    """
    search_dirs = [
        os.getcwd(),
        os.path.join(os.getcwd(), "Editable (Add CV here)")
    ]
    candidates_template = []
    candidates_general = []

    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        for name in os.listdir(d):
            low = name.lower()
            if not low.endswith(".docx"):
                continue
            if is_generated_output_name(low):
                continue
            full = os.path.join(d, name)
            if looks_like_template_name(low):
                candidates_template.append(full)
            else:
                candidates_general.append(full)

    pool = candidates_template or candidates_general
    if not pool:
        return None
    # newest by modification time
    return max(pool, key=os.path.getmtime)


def main():
    # Accept an optional CLI arg. If not provided, auto-find a .docx.
    in_path = sys.argv[1] if len(sys.argv) >= 2 else auto_find_input()

    if not in_path:
        print("Usage: python cv_splitter_v2.py <input_cv.docx>")
        print("No .docx found in the current folder or in 'Editable (Add CV here)'.")
        sys.exit(2)

    if not os.path.isfile(in_path):
        print(f"ERROR: File not found: {in_path}")
        sys.exit(2)

    print(f"Input CV: {os.path.abspath(in_path)}")

    person = parse_person_name_from_filename(in_path)
    out_full_docx = f"CenExel CURRICULUM VITAE {person}.docx"
    out_abbr_docx = f"CenExel Abbrv CURRICULUM VITAE {person}.docx"

    # Load original twice so each output keeps its section/header/footer relationships
    doc_abbr = Document(in_path)
    doc_full = Document(in_path)

    idx_abbr = find_first_signature_table_index(doc_abbr)
    idx_full = find_first_signature_table_index(doc_full)

    if idx_abbr is None or idx_full is None:
        print("ERROR: Could not locate the first Signature / Date of Signature box.")
        sys.exit(1)

    # Abbreviated: keep through the signature table
    remove_suffix_after(doc_abbr, idx_abbr)

    # Full: keep content AFTER the signature table
    remove_prefix(doc_full, idx_full + 1)

    # Fix 1: headers/footers visible on first page of both outputs
    disable_different_first_page(doc_abbr)
    disable_different_first_page(doc_full)

    # Fix 2: remove blank/page-break-only leading paragraphs in Full
    strip_leading_blank_or_pagebreak_paragraphs(doc_full)

    # Save DOCX first (required before conversion)
    doc_abbr.save(out_abbr_docx)
    doc_full.save(out_full_docx)

    # Convert to PDF
    out_abbr_pdf = out_abbr_docx.replace(".docx", ".pdf")
    out_full_pdf = out_full_docx.replace(".docx", ".pdf")

    if convert is None:
        print("WARNING: docx2pdf not available. Leaving .docx outputs in place.")
        print(f"  Abbreviated CV (DOCX): {os.path.abspath(out_abbr_docx)}")
        print(f"  Full CV        (DOCX): {os.path.abspath(out_full_docx)}")
        sys.exit(0)

    try:
        convert(out_abbr_docx, out_abbr_pdf)
        convert(out_full_docx, out_full_pdf)
        # Remove .docx if conversion succeeded
        try:
            os.remove(out_abbr_docx)
        except Exception:
            pass
        try:
            os.remove(out_full_docx)
        except Exception:
            pass

        print("Converted to PDF successfully.")
        print(f"  Abbreviated CV (PDF): {os.path.abspath(out_abbr_pdf)}")
        print(f"  Full CV        (PDF): {os.path.abspath(out_full_pdf)}")

    except Exception as e:
        print("PDF conversion failed:", e)
        print("Keeping .docx files:")
        print(f"  Abbreviated CV (DOCX): {os.path.abspath(out_abbr_docx)}")
        print(f"  Full CV        (DOCX): {os.path.abspath(out_full_docx)}")


if __name__ == '__main__':
    main()
