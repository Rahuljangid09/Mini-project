import pdfplumber
import re

# ── New ledger format (2019-course, cumulative multi-semester) ──────────────
# Student blocks are split on "SEAT NO.:" (old 8-digit+letter pattern no
# longer applies — seat numbers now look like B400230391).
SEAT_BLOCK_RE = re.compile(r'SEAT NO\.:\s*(\S+)')

# Header line: "SEAT NO.: B400230391 NAME : AADARSH RAJESH GORLE MOTHER : ..."
HEADER_RE = re.compile(r'SEAT NO\.:\s*(\S+)\s+NAME\s*:\s*(.+?)\s+MOTHER\s*:')

# Subject line: "404181 RADIATION & MICROWAVE THEORY 020/030 043/070 ... --- ---"
# Code first, subject NAME in the middle (variable length), then exactly
# 14 value tokens at the end: ISE ESE TOTAL TW PR OR TUT Tot% Crd Grd GP CP P&R ORD
SUBJECT_LINE_RE = re.compile(r'^(\d{6}[A-Z]?)\s+(.+)$')
N_VALUE_TOKENS = 14


# ── TEXT EXTRACTION ──────────────────────────────────────────────────────────

def extract_text_from_pdf(pdf_path: str) -> str:
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text


# ── MAIN PARSER ──────────────────────────────────────────────────────────────

def parse_ledger(raw_text: str) -> list:
    students = []
    matches = list(SEAT_BLOCK_RE.finditer(raw_text))

    for i, match in enumerate(matches):
        seat = match.group(1)
        block_start = match.start()
        block_end   = matches[i + 1].start() if i + 1 < len(matches) else len(raw_text)
        block = raw_text[block_start:block_end]

        name     = _extract_name(block)
        cgpa     = _extract_cgpa(block)
        subjects = _extract_all_subjects(block)
        status   = _get_status(subjects)

        students.append({
            "seat_no":  seat,
            "name":     name,
            "sgpa":     cgpa,   # cumulative ledger: CGPA is the headline figure
            "subjects": subjects,
            "status":   status,
        })

    return students


# ── FIELD HELPERS ─────────────────────────────────────────────────────────────

def _extract_name(block: str) -> str:
    m = HEADER_RE.search(block)
    return m.group(2).strip() if m else ""


def _extract_cgpa(block: str) -> str:
    m = re.search(r'CGPA\s*:\s*([\d.]+|--)', block)
    return m.group(1) if m else ""


# ── SUBJECT PARSING ───────────────────────────────────────────────────────────

def _parse_subject_line(line: str) -> dict | None:
    """
    Line = CODE  SUBJECT NAME (variable words, may include '*')  14 value tokens.

    Value order: ISE ESE TOTAL(raw) TW PR OR TUT Tot% Crd Grd GP CP P&R ORD
                  0   1     2        3  4  5  6   7    8   9   10 11  12  13
    """
    m = SUBJECT_LINE_RE.match(line.strip())
    if not m:
        return None

    merge_key = m.group(1)[:6]
    tokens = m.group(2).split()
    if len(tokens) < N_VALUE_TOKENS:
        return None

    values = tokens[-N_VALUE_TOKENS:]

    return {
        "merge_key": merge_key,
        "IN":    values[0],   # ISE
        "TH":    values[1],   # ESE
        "TW":    values[3],
        "PR":    values[4],
        "OR":    values[5],
        "TUT":   values[6],
        "TOTAL": values[7],   # Tot% (percentage)
        "GRADE": values[9],
    }


def _merge_subjects(existing: dict, new: dict) -> dict:
    """Fill missing fields from a second row of the same subject code
    (e.g. a subject examined across two semesters)."""
    for field in ["IN", "TH", "TW", "PR", "OR", "TUT", "TOTAL", "GRADE"]:
        if existing.get(field, "---") in ("---", "") \
                and new.get(field, "---") not in ("---", ""):
            existing[field] = new[field]
    return existing


def _extract_all_subjects(block: str) -> list:
    """Extracts every unique subject in the student's block — no count cap."""
    subj_map   = {}
    subj_order = []

    for line in block.split("\n"):
        s = _parse_subject_line(line)
        if not s:
            continue
        k = s["merge_key"]
        if k not in subj_map:
            subj_map[k] = s
            subj_order.append(k)
        else:
            subj_map[k] = _merge_subjects(subj_map[k], s)

    return [subj_map[k] for k in subj_order]


# ── STATUS LOGIC ──────────────────────────────────────────────────────────────

def _get_status(subjects: list) -> str:
    """
    - Any grade == F  → Fail
    - Any field contains 'AB' → AB
    - Otherwise → Pass
    """
    for s in subjects:
        if s.get("GRADE") == "F":
            return "Fail"
    for s in subjects:
        if any("AB" in s.get(f, "") for f in ["IN", "TH", "TW", "PR", "OR", "TUT", "TOTAL"]):
            return "AB"
    return "Pass"