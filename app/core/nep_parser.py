import pdfplumber
import re
from collections import defaultdict

# ── Fixed, document-wide column grid (confirmed identical across every
# page of the ledger — this is a template, not derived per page) ────────────

SLOT_X = [182.6, 222.7, 262.8, 302.8, 342.9, 383.0, 423.0, 463.1, 503.2, 543.2, 583.3, 623.4, 663.4]
# ISE/ESE repeat at different slots because different-credit subjects use
# different max-marks pairs (e.g. slot0/1 = 30/70 marks theory, slot6/7 =
# 15/35 marks a lighter theory course). Mapped to our existing internal
# field names (IN=ISE, TH=ESE) so excel_writer.py needs no changes.
SLOT_FIELDS = ["IN", "TH", "PR", "OR", "PR", "TW", "IN", "TH", "TW", "IN", "TUT", "OR", "IN"]

TOT_X = 692.1     # Total % — always numeric in this format, unlike older ledgers' "FF"
GRADE_X = 800.8   # Letter grade — confirmed empirically stable across all rows/pages
X_TOL = 4         # tolerance for matching a word's x0 to a known column position

HEADER_RE = re.compile(r"PRN:\s*(\S+)\s+SEAT NO\.:\s*(\S+)\s+NAME:\s*(.+?)\s+Mother's Name")

# Row-leading tokens that are never subject codes (page furniture, summary lines)
_NON_SUBJECT_KEYWORDS = {"SEMESTER:", "MEDIUM", "SECOND", "THIRD", "FOURTH", "FIRST", "PRN:"}


def _is_subject_row(first_word: dict) -> bool:
    """A subject code row: leftmost column, contains a digit, not a known keyword."""
    if abs(first_word["x0"] - 56.7) > 6:
        return False
    text = first_word["text"]
    if text in _NON_SUBJECT_KEYWORDS:
        return False
    return bool(re.search(r"\d", text))


def _nearest_word(words_by_x: list[tuple[float, str]], target_x: float, tol: float = X_TOL):
    for x0, text in words_by_x:
        if abs(x0 - target_x) <= tol:
            return text
    return None


def _group_rows(page):
    """Groups a page's words by visual row (rounded top), each row sorted left-to-right."""
    rows = defaultdict(list)
    for w in page.extract_words(use_text_flow=False, keep_blank_chars=False):
        rows[round(w["top"], 1)].append(w)
    return {top: sorted(words, key=lambda w: w["x0"]) for top, words in sorted(rows.items())}


def _parse_subject_row(line: list[dict]) -> dict | None:
    code = line[0]["text"]
    words_by_x = [(w["x0"], w["text"]) for w in line]

    fields: dict = {}
    for x, field in zip(SLOT_X, SLOT_FIELDS):
        val = _nearest_word(words_by_x, x)
        if val and val != "---":
            fields[field] = val

    total = _nearest_word(words_by_x, TOT_X)
    grade = _nearest_word(words_by_x, GRADE_X)

    return {
        "merge_key": code,   # no variant-suffix stripping needed — codes are already distinct (e.g. PCC-201-ETC)
        "subject_code": code,
        **fields,
        "TOTAL": total or "",
        "GRADE": grade or "",
    }


def parse_ledger_nep(pdf_path: str) -> list[dict]:
    """
    Parses the 2024 NEP-pattern ledger. Unlike the 2019-pattern parser,
    this one needs word position (x/y coordinates) from the PDF directly
    — plain extracted text is not reliable for this format's shared
    column-grid layout — so it takes a file path, not raw text.
    """
    students: list[dict] = []
    current: dict | None = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            rows = _group_rows(page)
            for top, line in rows.items():
                if not line:
                    continue
                joined = " ".join(w["text"] for w in line)

                header_match = HEADER_RE.search(joined)
                if header_match:
                    if current:
                        students.append(current)
                    prn_no, seat_no, name = header_match.groups()
                    current = {
                        "prn_no": prn_no,
                        "seat_no": seat_no,
                        "name": name.strip(),
                        "subjects": [],
                    }
                    continue

                if current is None:
                    continue  # page furniture before the first student

                if _is_subject_row(line[0]):
                    subj = _parse_subject_row(line)
                    if subj:
                        current["subjects"].append(subj)

    if current:
        students.append(current)

    for s in students:
        s["status"] = _get_status(s["subjects"])

    return students


def _get_status(subjects: list[dict]) -> str:
    for s in subjects:
        if s.get("GRADE") == "F":
            return "Fail"
    for s in subjects:
        if any("AB" in s.get(f, "") for f in ["IN", "TH", "TW", "PR", "OR", "TUT", "TOTAL"]):
            return "AB"
    return "Pass"
