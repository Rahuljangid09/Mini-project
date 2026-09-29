"""
Detects which ledger format a PDF is, so main.py can dispatch to the
right parser automatically — no user-facing format selector needed.

Adding support for a future new SPPU pattern means adding one new
parser module + one new elif here — nothing else in the pipeline
(excel_writer.py, main.py's routes) needs to change, since every
parser produces the same internal student/subject shape.
"""

import pdfplumber

FORMAT_2019 = "2019_PATTERN"
FORMAT_2024_NEP = "2024_NEP"


def detect_format(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""

    if "B.E.(2019 COURSE)" in first_page_text:
        return FORMAT_2019
    if "S.E. (2024 Pattern (NEP 2020)" in first_page_text:
        return FORMAT_2024_NEP

    raise ValueError(
        "Unrecognized ledger format — this PDF doesn't match any known "
        "pattern (2019-course or 2024 NEP). A new parser module is needed."
    )
