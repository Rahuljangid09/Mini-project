import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Constants ────────────────────────────────────────────────────────────────

VALID_GRADES = {"O", "A+", "A", "B+", "B", "P"}

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
        counts = {g: 0 for g in ["O", "A+", "A", "B+", "B", "P"]}
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


# ─── Main Entry Point ─────────────────────────────────────────────────────────

def generate_excel(all_students: list) -> str:
    """
    Each student → one row.
    Columns: Seat No, Student Name,
             <CODE1>_IN, <CODE1>_TH, <CODE1>_TOTAL, <CODE1>_GRADE,
             ... (up to 5 subjects) ...
             Final Status
    """

    # ── Collect all subject codes in a stable order ──────────────────────────
    all_codes = []
    seen_codes = set()
    for student in all_students:
        for subj in student.get("subjects", []):
            code = _subject_code(subj)
            if code and code not in seen_codes:
                seen_codes.add(code)
                all_codes.append(code)
        if len(all_codes) >= 5:
            break
    all_codes = all_codes[:5]

    # ── Build rows ────────────────────────────────────────────────────────────
    rows = []
    for student in all_students:
        row = {
            "Seat No":      student.get("seat_no", ""),
            "Student Name": student.get("name", ""),
        }
        for code in all_codes:
            row[f"{code}_IN"]    = ""
            row[f"{code}_TH"]    = ""
            row[f"{code}_TOTAL"] = ""
            row[f"{code}_GRADE"] = ""

        for subj in student.get("subjects", []):
            code = _subject_code(subj)
            if code not in seen_codes:
                continue
            row[f"{code}_IN"]    = subj.get("IN",    "")
            row[f"{code}_TH"]    = subj.get("TH",    "")
            row[f"{code}_TOTAL"] = subj.get("TOTAL", "")
            row[f"{code}_GRADE"] = subj.get("GRADE", "")

        row["Final Status"] = student.get("status", "")
        rows.append(row)

    # ── Enforce strict column order ───────────────────────────────────────────
    fixed_cols   = ["Seat No", "Student Name"]
    subject_cols = []
    for code in all_codes:
        subject_cols += [f"{code}_IN", f"{code}_TH", f"{code}_TOTAL", f"{code}_GRADE"]
    all_cols = fixed_cols + subject_cols + ["Final Status"]

    df = pd.DataFrame(rows)
    for col in all_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[all_cols]

    # ── Save Sheet 1 ──────────────────────────────────────────────────────────
    os.makedirs("outputs", exist_ok=True)
    path = "outputs/result_analysis.xlsx"

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Raw Data")

    # ── Add Sheet 2 ───────────────────────────────────────────────────────────
    _add_analysis_sheet(path, all_students, all_codes)

    return path
