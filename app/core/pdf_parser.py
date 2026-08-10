import pdfplumber
import re

# Matches seat number only at the START of a line (avoids duplicate mid-line match)
SEAT_LINE_RE = re.compile(r'(?:^|\n)(\d{8}[A-Z])\b', re.MULTILINE)

# Matches subject rows: CODE * tokens...
SUBJECT_RE = re.compile(r'^(\d{6}[A-Z]?)\s+\*\s+(.+)$')


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
    matches = list(SEAT_LINE_RE.finditer(raw_text))

    for i, match in enumerate(matches):
        seat = match.group(1)

        # Block starts right at this seat line, ends at next seat line
        block_start = match.start() if match.start() == 0 else match.start() + 1
        block_end   = matches[i + 1].start() if i + 1 < len(matches) else len(raw_text)
        block = raw_text[block_start:block_end].strip()

        lines = block.split("\n")

        name     = _extract_name(lines[0])
        sgpa     = _extract_sgpa(block)
        subjects = _extract_first5_subjects(lines[1:])
        status   = _get_status(subjects)

        students.append({
            "seat_no":  seat,
            "name":     name,
            "sgpa":     sgpa,
            "subjects": subjects,
            "status":   status,
        })

    return students


# ── FIELD HELPERS ─────────────────────────────────────────────────────────────

def _extract_name(header_line: str) -> str:
    """
    Header line looks like:
      71608229C  NIKAM YASHODEEP PANDHARINATH SHOBHA  71608229C SCEP
    Name sits between the two seat occurrences.
    """
    seats = re.findall(r'\d{8}[A-Z]', header_line)
    if len(seats) >= 2:
        idx1 = header_line.index(seats[0]) + len(seats[0])
        idx2 = header_line.index(seats[1], idx1)
        return header_line[idx1:idx2].strip()
    # Fallback: everything after the seat number
    parts = header_line.split()
    return " ".join(parts[1:]) if len(parts) > 1 else ""


def _extract_sgpa(block: str) -> str:
    m = re.search(r'SGPA\d*\s*:\s*([\d.]+|--)', block)
    return m.group(1) if m else ""


# ── SUBJECT PARSING ───────────────────────────────────────────────────────────

def _parse_subject_line(line: str) -> dict | None:
    """
    Subject line format (13 tokens after removing *):
      CODE * IN  TH  IN+TH  TW  PR  OR  Tot%  Crd  Grd  GP  CP  P&R  ORD
       idx:  0   1    2      3   4   5    6     7    8    9  10  11   12

    We extract only: IN(0), TH(1), Tot%(6), Grd(8)
    """
    m = SUBJECT_RE.match(line.strip())
    if not m:
        return None

    # Normalize code: strip variant letter suffix → merge key
    # e.g. 404184A → 404184, 404185B → 404185
    merge_key = m.group(1)[:6]

    tokens = m.group(2).split()
    if len(tokens) < 9:
        return None

    return {
        "merge_key": merge_key,
        "IN":    tokens[0],   # Internal marks  e.g. 017/030
        "TH":    tokens[1],   # Theory marks    e.g. 034/070
        "TOTAL": tokens[6],   # Tot%            e.g. 51 or FF
        "GRADE": tokens[8],   # Grade           e.g. B, A+, O, F
    }


def _merge_subjects(existing: dict, new: dict) -> dict:
    """Fill missing fields from the second row of the same subject code."""
    for field in ["IN", "TH", "TOTAL", "GRADE"]:
        if existing.get(field, "---") in ("---", "") \
                and new.get(field, "---") not in ("---", ""):
            existing[field] = new[field]
    return existing


def _extract_first5_subjects(lines: list) -> list:
    subj_map   = {}   # merge_key → subject dict
    subj_order = []   # preserve first-seen order

    for line in lines:
        s = _parse_subject_line(line)
        if not s:
            continue
        k = s["merge_key"]
        if k not in subj_map:
            subj_map[k] = s
            subj_order.append(k)
        else:
            subj_map[k] = _merge_subjects(subj_map[k], s)

    # Return only the first 5 unique subjects
    # NOTE: this [:5] cap is exactly what Issue 3 (all-subjects) removes next.
    return [subj_map[k] for k in subj_order[:5]]


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
        if any("AB" in s.get(f, "") for f in ["IN", "TH", "TOTAL"]):
            return "AB"
    return "Pass"
