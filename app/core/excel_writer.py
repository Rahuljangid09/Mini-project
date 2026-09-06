import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Constants ────────────────────────────────────────────────────────────────

VALID_GRADES = {"O", "A+", "A", "B+", "B", "C", "P"}

HEADER_FILL = PatternFill("solid", start_color="2F5496", end_color="2F5496")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
TITLE_FONT  = Font(bold=True, name="Arial", size=11, color="1F3864")
CELL_FONT   = Font(name="Arial", size=10)
CENTER      = Alignment(horizontal="center", vertical="center")
LEFT        = Alignment(horizontal="left",   vertical="center")
ALT_FILL    = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)


# ─── Analysis Helpers ─────────────────────────────────────────────────────────

def _safe_pct(numerator, denominator):
    if denominator == 0:
        return 0.0
    return round(numerator / denominator * 100, 2)

def _is_ab(value):
    return str(value).strip().upper().startswith("AB")

def _is_fail(grade):
    return str(grade).strip().upper() == "F"

def _subject_code(sub: dict) -> str:
    """Single source of truth for reading a subject's code (merge_key)."""
    return sub.get("merge_key", "")

def _style_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER

def _style_data(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        fill = ALT_FILL if (r - r1) % 2 == 1 else None
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = CELL_FONT
            cell.alignment = CENTER
            cell.border = THIN_BORDER
            if fill:
                cell.fill = fill

def _write_table(ws, df, title, start_row, start_col=1):
    ws.cell(row=start_row, column=start_col, value=title).font = TITLE_FONT

    header_row = start_row + 1
    for i, col in enumerate(df.columns, start=start_col):
        ws.cell(row=header_row, column=i, value=col)
    _style_header(ws, header_row, start_col, start_col + len(df.columns) - 1)

    data_start = header_row + 1
    for r_off, (_, row) in enumerate(df.iterrows()):
        for i, val in enumerate(row, start=start_col):
            ws.cell(row=data_start + r_off, column=i, value=val)
    _style_data(ws, data_start, data_start + len(df) - 1, start_col, start_col + len(df.columns) - 1)

    for i, col in enumerate(df.columns, start=start_col):
        max_len = max(len(str(col)), *(len(str(v)) for v in df.iloc[:, i - start_col]))
        ws.column_dimensions[get_column_letter(i)].width = max_len + 4

    return data_start + len(df)


# ─── Table Builders ───────────────────────────────────────────────────────────

def _table1(all_students, all_codes):
    rows = []
    for code in all_codes:
        total = pass_ = fail = ab = 0
        for s in all_students:
            for sub in s.get("subjects", []):
                if _subject_code(sub) != code:
                    continue
                grade  = str(sub.get("GRADE", "")).strip().upper()
                in_val = str(sub.get("IN",    "")).strip().upper()
                th_val = str(sub.get("TH",    "")).strip().upper()
                if _is_ab(in_val) or _is_ab(th_val) or grade in ("AB", "---", ""):
                    ab += 1
                elif _is_fail(grade):
                    fail += 1
                else:
                    pass_ += 1
                total += 1
        present = total - ab
        rows.append({
            "Subject": code,
            "Total Students": total,
            "Pass": pass_,
            "Fail": fail,
            "AB": ab,
            "Present Students": present,
            "Pass %": _safe_pct(pass_, total),
            "Pass % (Without AB)": _safe_pct(pass_, present),
        })
    return pd.DataFrame(rows)


def _table2(all_students, all_codes):
    rows = []
    for code in all_codes:
        counts = {g: 0 for g in ["O", "A+", "A", "B+", "B", "C", "P"]}
        for s in all_students:
            for sub in s.get("subjects", []):
                if _subject_code(sub) != code:
                    continue
                g = str(sub.get("GRADE", "")).strip()
                if g in VALID_GRADES:
                    counts[g] += 1
        rows.append({"Sub Code": code, **counts})
    return pd.DataFrame(rows)


def _table3(all_students):
    total     = len(all_students)
    all_clear = sum(1 for s in all_students if not any(_is_fail(sub.get("GRADE", "")) for sub in s.get("subjects", [])))
    fail      = total - all_clear
    return pd.DataFrame([
        {"Metric": "Total Students", "Value": total},
        {"Metric": "All Clear",      "Value": all_clear},
        {"Metric": "Fail",           "Value": fail},
        {"Metric": "% All Clear",    "Value": _safe_pct(all_clear, total)},
    ])


def _table4(all_students):
    present   = [s for s in all_students if str(s.get("status", "")).strip().upper() != "AB"]
    total     = len(present)
    all_clear = sum(1 for s in present if not any(_is_fail(sub.get("GRADE", "")) for sub in s.get("subjects", [])))
    fail      = total - all_clear
    return pd.DataFrame([
        {"Metric": "Total Students (excl. AB)", "Value": total},
        {"Metric": "All Clear",                 "Value": all_clear},
        {"Metric": "Fail",                      "Value": fail},
        {"Metric": "% All Clear",               "Value": _safe_pct(all_clear, total)},
    ])


def _add_analysis_sheet(path, all_students, all_codes):
    wb = load_workbook(path)
    if "Analysis" in wb.sheetnames:
        del wb["Analysis"]
    ws = wb.create_sheet("Analysis")

    current_row = 1
    for title, df in [
        ("Subject-wise Summary",            _table1(all_students, all_codes)),
        ("Subject-wise Grade Distribution", _table2(all_students, all_codes)),
        ("Overall Summary (With AB)",       _table3(all_students)),
        ("Overall Summary (Without AB)",    _table4(all_students)),
    ]:
        current_row = _write_table(ws, df, title, current_row) + 2

    wb.save(path)
    wb.close()


# ─── Per-subject dynamic sub-columns ──────────────────────────────────────────

# (display label, source field in the parsed subject dict)
_POSSIBLE_SUBCOLS = [
    ("ISE",   "IN"),
    ("ESE",   "TH"),
    ("TOTAL", "TOTAL"),
    ("TV",    "TW"),
    ("PR",    "PR"),
    ("OR",    "OR"),
]

def _blank(v) -> bool:
    v = str(v).strip()
    return v in ("", "---")

def _active_subcolumns(all_students, code):
    """
    Only include a sub-column (ISE/ESE/TOTAL/TV/PR/OR) for this subject
    code if at least one student has real data in it. GRADE is always
    included. This is what makes theory subjects show ISE/ESE/TOTAL and
    lab/project subjects show only TV/PR/OR as applicable — no manual
    column deletion needed.
    """
    active = []
    for label, field in _POSSIBLE_SUBCOLS:
        has_data = any(
            _subject_code(sub) == code and not _blank(sub.get(field, ""))
            for s in all_students
            for sub in s.get("subjects", [])
        )
        if has_data:
            active.append((label, field))
    active.append(("GRADE", "GRADE"))
    return active


def _style_two_row_header(ws, n_cols):
    for row in (1, 2):
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = CENTER
            cell.border = THIN_BORDER


def _style_raw_data_rows(ws, n_rows, n_cols):
    for r in range(3, n_rows + 3):
        fill = ALT_FILL if (r - 3) % 2 == 1 else None
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = CELL_FONT
            cell.alignment = CENTER
            cell.border = THIN_BORDER
            if fill:
                cell.fill = fill
        ws.row_dimensions[r].height = 16


def _autofit_columns(ws, n_cols):
    for c in range(1, n_cols + 1):
        col_letter = get_column_letter(c)
        max_len = 0
        for cell in ws[col_letter]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 3


# ─── Main Entry Point ─────────────────────────────────────────────────────────

def generate_excel(all_students: list) -> str:
    """
    Each student → one row. Every subject present in the ledger is
    included (no 5-subject cap). Each subject gets only the sub-columns
    it actually has data for (e.g. theory: ISE/ESE/TOTAL/GRADE, lab: TV/GRADE,
    lab with oral: TV/OR/GRADE), grouped under a merged subject-code header.
    """

    # ── Collect every subject code, in first-seen order, no cap ──────────────
    all_codes = []
    seen_codes = set()
    for student in all_students:
        for subj in student.get("subjects", []):
            code = _subject_code(subj)
            if code and code not in seen_codes:
                seen_codes.add(code)
                all_codes.append(code)

    # ── Decide active sub-columns per subject ─────────────────────────────────
    subject_columns = []  # list of (code, label, field)
    for code in all_codes:
        for label, field in _active_subcolumns(all_students, code):
            subject_columns.append((code, label, field))

    # ── Build two-level column headers (merged cells via pandas MultiIndex) ──
    col_tuples = [("Seat No", ""), ("Student Name", "")]
    col_tuples += [(code, label) for code, label, _ in subject_columns]
    col_tuples += [("Final Status", "")]

    # ── Build rows ────────────────────────────────────────────────────────────
    rows = []
    for student in all_students:
        subj_by_code = {_subject_code(s): s for s in student.get("subjects", [])}
        row = [student.get("seat_no", ""), student.get("name", "")]
        for code, label, field in subject_columns:
            subj = subj_by_code.get(code)
            row.append(subj.get(field, "") if subj else "")
        row.append(student.get("status", ""))
        rows.append(row)

    # ── Save Sheet 1 (written directly via openpyxl for merged 2-row header) ──
    os.makedirs("outputs", exist_ok=True)
    path = "outputs/result_analysis.xlsx"

    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Raw Data"

    n_cols = len(col_tuples)

    # Row 1: top-level labels, merging consecutive cells with the same code
    # Row 2: sub-column labels (blank for the fixed Seat No / Name / Status cols)
    c = 1
    while c <= n_cols:
        top_label = col_tuples[c - 1][0]
        span = 1
        while c + span - 1 < n_cols and col_tuples[c + span - 1][0] == top_label \
                and top_label not in ("Seat No", "Student Name", "Final Status"):
            span += 1
        ws.cell(row=1, column=c, value=top_label)
        if span > 1:
            ws.merge_cells(start_row=1, start_column=c, end_row=1, end_column=c + span - 1)
        else:
            ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
        c += span

    for i, (top, sub) in enumerate(col_tuples, start=1):
        if top not in ("Seat No", "Student Name", "Final Status"):
            ws.cell(row=2, column=i, value=sub)

    # Data rows
    for r_off, row_data in enumerate(rows, start=3):
        for i, val in enumerate(row_data, start=1):
            ws.cell(row=r_off, column=i, value=val)

    n_rows = len(rows)
    _style_two_row_header(ws, n_cols)
    _style_raw_data_rows(ws, n_rows, n_cols)
    _autofit_columns(ws, n_cols)

    wb.save(path)

    # ── Add Sheet 2 (now covers every subject, not just the first 5) ─────────
    _add_analysis_sheet(path, all_students, all_codes)

    return path