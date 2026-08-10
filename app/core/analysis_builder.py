"""
Builds the dashboard's analysis_data dict from parsed student records.
Moved out of main.py so routes stay thin and this logic is unit-testable
on its own.
"""


def _subject_code(sub: dict) -> str:
    """Single source of truth for reading a subject's code (merge_key)."""
    return sub.get("merge_key", "")


def build_analysis_data(all_students: list) -> dict:
    """Build the dashboard analysis_data dict directly from all_students."""

    total          = len(all_students)
    ab_students    = sum(1 for s in all_students if str(s.get("status", "")).strip().upper() == "AB")
    fail_students  = sum(1 for s in all_students if str(s.get("status", "")).strip().upper() == "FAIL")
    pass_students  = total - ab_students - fail_students
    pass_pct       = round(pass_students / total * 100, 1) if total else 0

    # Subject stats
    subject_map = {}
    for s in all_students:
        for sub in s.get("subjects", []):
            code  = _subject_code(sub)
            grade = str(sub.get("GRADE", "")).strip().upper()
            in_v  = str(sub.get("IN",    "")).strip().upper()
            th_v  = str(sub.get("TH",    "")).strip().upper()
            if code not in subject_map:
                subject_map[code] = {"code": code, "pass": 0, "fail": 0, "ab": 0}
            if in_v.startswith("AB") or th_v.startswith("AB") or grade in ("AB", "---", ""):
                subject_map[code]["ab"] += 1
            elif grade == "F":
                subject_map[code]["fail"] += 1
            else:
                subject_map[code]["pass"] += 1

    subjects = []
    for code, d in subject_map.items():
        total_sub = d["pass"] + d["fail"] + d["ab"]
        present   = total_sub - d["ab"]
        d["pass_percent"] = round(d["pass"] / present * 100, 1) if present else 0
        subjects.append(d)

    # Grade distribution
    grade_counts = {"O": 0, "A+": 0, "A": 0, "B+": 0, "B": 0, "P": 0, "F": 0}
    for s in all_students:
        for sub in s.get("subjects", []):
            g = str(sub.get("GRADE", "")).strip()
            if g in grade_counts:
                grade_counts[g] += 1

    return {
        "summary": {
            "total_students": total,
            "pass":           pass_students,
            "fail":           fail_students,
            "ab":             ab_students,
            "pass_percent":   pass_pct,
        },
        "subjects": subjects,
        "grades":   grade_counts,
    }
