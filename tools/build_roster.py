"""Build roster.html from data/members.xlsx + tools/analysis.py.

Run from anywhere:  python tools/build_roster.py
Needs: pip install openpyxl
"""
import json
import re
from pathlib import Path
from urllib.parse import quote

import openpyxl

from analysis import ANALYSIS, INCOMPLETE_NOTES

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "members.xlsx"
CV_DIR = ROOT / "data" / "cv"
TEMPLATE = Path(__file__).with_name("roster_template.html")
OUT = ROOT / "roster.html"
LEAD_EMAIL = "shumokhalsharif@gmail.com"  # track lead: shown in the header, not as a member card

MAJORS = {
    "هندسة برمجيات": "Software Engineering",
    "علوم حاسب": "Computer Science",
    "هندسة حاسب": "Computer Engineering",
    "ذكاء اصطناعي": "Artificial Intelligence",
    "هندسة كهربائية": "Electrical Engineering",
    "إعلام": "Media",
}
FIT = {"مطابق": "Match", "قريب": "Related", "لا": "Not related"}
YEARS = {
    "السنة الأولى": 1, "السنة الثانية": 2, "السنة الثالثة": 3,
    "السنة الرابعة": 4, "السنة الخامسة": 5,
}


def clean(v):
    if v is None:
        return None
    s = str(v).strip()
    return None if s.lower() in ("", "none") else s


def phone(v):
    s = clean(v)
    if not s:
        return None
    s = re.sub(r"\D", "", s)
    return s if s.startswith("0") else "0" + s


def local_cvs():
    by_id = {}
    for f in CV_DIR.iterdir():
        m = re.search(r"(\d{7})", f.stem)
        if m:
            by_id[m.group(1)] = f
    return by_id


def main():
    ws = openpyxl.load_workbook(XLSX).active
    cvs = local_cvs()
    members = []
    lead = None
    for row in ws.iter_rows():
        name = clean(row[0].value)
        if not name:
            continue
        email = clean(row[6].value)
        if (email or "").lower() == LEAD_EMAIL:
            lead = name
            continue
        cv_url = row[7].hyperlink.target if row[7].hyperlink else None
        pf_url = row[8].hyperlink.target if row[8].hyperlink else None
        sid = None
        if cv_url:
            m = re.search(r"_(\d{7})@", cv_url)
            sid = m.group(1) if m else None
        local = cvs.get(sid) if sid else None
        gpa = row[4].value
        key = (email or "").lower()
        a = ANALYSIS.get(key)

        notes = list(a["notes"]) if a else []
        if not a:
            if key in INCOMPLETE_NOTES:
                notes.insert(0, INCOMPLETE_NOTES[key])
            elif not cv_url:
                notes.insert(0, "No CV submitted.")
            elif not local:
                notes.insert(0, "CV is linked but the file is not in data/cv.")

        major = clean(row[1].value)
        members.append({
            "name": name,
            "en": a["en"] if a else None,
            "major": MAJORS.get(major, major),
            "fit": FIT.get(clean(row[2].value)),
            "year": YEARS.get(clean(row[3].value)),
            "gpa": gpa if isinstance(gpa, (int, float)) else None,
            "phone": phone(row[5].value),
            "email": email,
            "cv": cv_url,
            "cvLocal": "data/cv/" + quote(local.name) if local else None,
            "portfolio": pf_url,
            "complete": bool(a),
            "level": a["level"] if a else None,
            "focus": a["focus"] if a else [],
            "skills": a["skills"] if a else [],
            "projects": a["projects"] if a else [],
            "experience": a["experience"] if a else [],
            "summary": a["summary"] if a else None,
            "strengths": a["strengths"] if a else None,
            "gaps": a["gaps"] if a else None,
            "next": a["next"] if a else None,
            "links": a["links"] if a else {},
            "notes": notes,
        })

    missing = set(ANALYSIS) - {(m["email"] or "").lower() for m in members}
    if missing:
        print("Analysis entries with no matching sheet row:", missing)

    data = json.dumps(members, ensure_ascii=False)
    OUT.write_text(TEMPLATE.read_text(encoding="utf-8").replace("/*DATA*/[]", data).replace('/*LEAD*/""', json.dumps(lead or "", ensure_ascii=False)), encoding="utf-8")
    done = sum(m["complete"] for m in members)
    print(f"Wrote {OUT.name}: {len(members)} members, {done} analyzed, {len(members) - done} incomplete")


if __name__ == "__main__":
    main()
