r"""Turn college catalog PDFs (program requirement pages) into a Power BI-ready Excel workbook.

Reads both layouts automatically: the classic two-column catalog (2020-21 to 2024-25) and the
CurriQunet META export (2025-26 onward). Each PDF's catalog year comes from its file name.

Usage:
    python catalog_to_excel.py "C:\path\to\Catalogs" -o requirements.xlsx
    python catalog_to_excel.py City_2023-2024_catalog_organized.pdf -o requirements_2023-24.xlsx
    python catalog_to_excel.py "C:\path\to\Catalogs" -o requirements.xlsx --no-description

A folder uses every PDF in it whose name contains a catalog year (e.g. 2023-2024); other PDFs are skipped.

Requires: pandas, openpyxl, and Poppler's pdftotext on PATH
(Windows: `conda install -c conda-forge poppler`, or download Poppler and add its bin folder to PATH).
"""
import argparse
import datetime as dt
import re
import subprocess
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

CODE = r"[A-Z]{2,5}\s+[A-Z]?[0-9]{1,4}[A-Z]{0,3}"
STATUS = r"(Active|Inactive|Approved|Pending)"
COURSE_RE = re.compile(
    r"^\s*([A-Z]{2,5})\s+([A-Z]?[0-9]{1,4}[A-Z]{0,3})\s{2,}(.+?)\s*(?:\(" + STATUS + r"\))?\s{2,}"
    r"([0-9]+(?:\.[0-9]+)?)(?:\s*-\s*([0-9]+(?:\.[0-9]+)?))?\s*$")
# Course lines right after a page break can lose their units column; units then come from the course description.
COURSE_NOUNITS_RE = re.compile(
    r"^\s*([A-Z]{2,5})\s+([A-Z]?[0-9]{1,4}[A-Z]{0,3})\s{2,}(.+?)\s*\(" + STATUS + r"\)\s*$")
UNITS_ONLY_RE = re.compile(r"^\s+([0-9]+\.[0-9]+)(?:\s*-\s*([0-9]+\.?[0-9]*))?\s*$")
BLOCK_RE = re.compile(r"^(\S.*?)\s{2,}(?:Semester )?Units:\s*([0-9.]+)(?:\s*-\s*([0-9.]+))?\s*$")
SUBBLOCK_RE = re.compile(
    r"^(?=\S)(.*?(?:select|complete|choose).*?)\s{2,}([0-9]+\.[0-9]+)(?:\s*-\s*([0-9]+\.?[0-9]*))?\s*$", re.I)
AREA_HEADER_RE = re.compile(r"^(\S.*?)\s{2,}(?:Semester )?Units:\s*(?:0(?:\.0)?)?\s*$")   # "ADMINISTRATION  Semester Units: 0.0"
SUBBLOCK_NOUNITS_RE = re.compile(r"^(?=\S)(.*?\b(?:select|complete|choose)\b.*?(?:following|from)\s*:?)\s*$", re.I)
TOTAL_RE = re.compile(r"Total(?: Units)?:\s*([0-9.]+)(?:\s*-\s*([0-9.]+))?")
HEADER_END_RE = re.compile(r":\s*CITY\s*$", re.I)
WRAPPED_TITLE_RE = re.compile(r"\(" + STATUS + r"\)\s*$")
AWARDS = [
    (r"ASSOCIATE IN ARTS FOR TRANSFER", "AA-T"), (r"ASSOCIATE IN SCIENCE FOR TRANSFER", "AS-T"),
    (r"ASSOCIATE OF ARTS", "AA"), (r"ASSOCIATE OF SCIENCE", "AS"),
    (r"CERTIFICATE OF ACHIEVEMENT", "Certificate of Achievement"),
    (r"CERTIFICATE OF PERFORMANCE", "Certificate of Performance"),
    (r"BACHELOR OF SCIENCE", "BS"), (r"LICENSURE ONLY", "Licensure only"),
]
NUM_WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
             "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12}


# ---------------------------------------------------------------- text extraction
def pdf_to_text(pdf, layout):
    args = ["pdftotext"] + (["-layout"] if layout else []) + [str(pdf), "-"]
    return subprocess.run(args, capture_output=True, check=True).stdout.decode("utf-8")


def clean_lines(text):
    lines = [l.rstrip() for l in text.replace("\f", "\n").splitlines()
             if not re.fullmatch(r"\s*\d{1,3}\s*", l)]          # drop page numbers
    joined = []
    for l in lines:                                              # "Semester Units: 40.0-" + "41.0"
        if joined and re.search(r"-\s*$", joined[-1]) and re.fullmatch(r"\s+[0-9]+\.?[0-9]*", l):
            joined[-1] = joined[-1].rstrip() + l.strip()
        else:
            joined.append(l)
    return joined


def description_units(plain_text):
    units = {}
    for m in re.finditer(r"^(" + CODE + r") [^\n]*\n[^\n]*?;\s*([0-9.]+)(?:\s*-\s*[0-9.]+)?\s+units?\s*$",
                         plain_text, re.M):
        units.setdefault(re.sub(r"\s+", " ", m.group(1)), float(m.group(2)))
    return units


# ---------------------------------------------------------------- parsing
def num(tok):
    tok = tok.lower()
    if tok in NUM_WORDS:
        return NUM_WORDS[tok]
    try:
        return float(tok)
    except ValueError:
        return None


def classify(header):
    """Return (rule_type, min_required, needs_review) for a requirement header."""
    h = re.sub(r"\s+", " ", header.lower().replace("–", "-"))
    top = re.search(r"if needed to (?:total|reach) (\d+(?:\.\d+)?) units", h)
    if top:                                                      # "If needed to total 18 units, select ..."
        return "TOPUP", float(top.group(1)), False
    review = bool(re.search(r"not (selected|already)|areas?\b|maximum|option|sequence|prerequisite|additional", h))
    lead = re.match(r"(\d+(?:\.\d+)?)(?:\s*(?:to|-)\s*\d+(?:\.\d+)?)?\s+(?:semester\s+)?units?\b", h)
    if lead:
        return "UNITS", float(lead.group(1)), review
    cu = re.search(r"(?:select|complete|choose)\b.*?\b(?:courses?|specializations?)\b.*?\((?:a\s+)?(?:minimum\s+(?:of\s+)?)?(\d+(?:\.\d+)?)"
                   r"(?:\s*-\s*\d+(?:\.\d+)?)?\s*(?:semester\s+)?units?", h)
    if cu and not re.search(r"\bor\b", h):                      # course count AND a unit minimum: units win
        return "UNITS", float(cu.group(1)), review
    pu = re.search(r"(?:select|complete|choose)\s+(?:a\s+)?(?:minimum\s+(?:of\s+)?)?\w+\s*\(\s*(\d+(?:\.\d+)?)\s*(?:semester\s+)?units?\s*\)", h)
    if pu:
        return "UNITS", float(pu.group(1)), review
    m = re.search(r"(?:select|complete|choose)\s+(?:a\s+)?(?:at\s+least\s+)?(?:minimum\s+(?:of\s+)?)?(\w+(?:\.\d+)?)"
                  r"(?:\s*(?:to|-)\s*\w+)?\s*(?:or more\s+)?(?:\([^)]*\)\s*)?(?:additional\s+)?(units?|courses?|of)", h)
    if not (m and num(m.group(1)) is not None):
        if re.search(r"\b(select|choose)\b", h):
            return None, None, True
        return "ALL", None, review
    n, kind = num(m.group(1)), m.group(2)
    return ("UNITS", float(n), review) if kind.startswith("unit") else ("COURSES", int(n), review)


def find_program_headers(lines):
    headers = []
    def next_text(i):
        return next((x.strip() for x in lines[i + 1:i + 4] if x.strip()), "")
    for i, l in enumerate(lines):
        bare = (not l.startswith(" ") and l.strip().isupper() and next_text(i) == "Summary"
                and re.search(r"(DEGREE|ACHIEVEMENT|PERFORMANCE|ONLY\))\s*$", l.strip()))   # heading without ": CITY"
        if (HEADER_END_RE.search(l) or bare) and not l.startswith(" "):
            start = i
            while start - 1 >= 0 and i - start < 3:              # program names can wrap onto 2-3 lines
                prev = lines[start - 1]
                if (not prev.strip() or prev.startswith(" ") or HEADER_END_RE.search(prev)
                        or re.search(r"[a-z]", prev) or "Total Units" in prev or prev.strip() == "Courses"):
                    break
                start -= 1
            headers.append((start, i, " ".join(x.strip() for x in lines[start:i + 1])))
    return headers


def parse_requirement_lines(req, desc_units):
    """Group requirement lines into blocks -> slots (OR alternatives) -> options (AND sequences) -> courses."""
    blocks, cur, slot, option, connector, area, courses = [], None, None, None, None, "", {}
    nonblank = [l for l in req if l.strip()]
    for j, l in enumerate(nonblank):
        s = l.strip()
        am = AREA_HEADER_RE.match(l)
        if am and cur is not None:
            area = am.group(1).strip().title()
            cur["notes"].append(f"[area] {area}")
            slot = option = connector = None
            continue
        cm = COURSE_RE.match(l)
        bm = None if cm else BLOCK_RE.match(l)
        sm = None if (cm or bm) else SUBBLOCK_RE.match(l)
        nm = None if (cm or bm or sm or cur is None or not cur["slots"]) else SUBBLOCK_NOUNITS_RE.match(l)
        if nm:                                                   # "Select one course from the following:" with no units
            cur = {"header": nm.group(1).strip(), "lo": 0.0, "hi": 0.0, "slots": [], "notes": [], "sub": True}
            blocks.append(cur)
            slot = option = connector = None
            area = ""
            continue
        if bm or sm:
            m = bm or sm
            cur = {"header": m.group(1).strip(), "lo": float(m.group(2)), "hi": float(m.group(3) or m.group(2)),
                   "slots": [], "notes": [], "sub": bool(sm)}
            blocks.append(cur)
            slot = option = connector = None
            area = ""
            continue
        if cur is None:
            continue
        fields = None
        if cm:
            fields = cm.groups()
        elif COURSE_NOUNITS_RE.match(l):
            subj, cnum, title, status = COURSE_NOUNITS_RE.match(l).groups()
            u = desc_units.get(f"{subj} {cnum}")
            if u is None:
                cur["notes"].append(f"{subj} {cnum}: units not found")
                u = 0.0
            fields = (subj, cnum, title, status, str(u), None)
        if fields:
            subj, cnum, title, _, u1, u2 = fields
            c = {"code": f"{subj} {cnum}", "units": float(u1), "units_max": float(u2 or u1), "area": area}
            courses.setdefault(c["code"], {"subject": subj, "title": title.strip(),
                                           "units": c["units"], "units_max": c["units_max"]})
            if connector == "AND" and option is not None:
                option.append(c)
            elif connector == "OR" and slot is not None:
                option = [c]
                slot.append(option)
            else:
                option = [c]
                slot = [option]
                cur["slots"].append(slot)
            connector = None
            continue
        if UNITS_ONLY_RE.match(l) and cur is not None:          # units line opens a new OR/AND cluster
            nxt = nonblank[j + 1] if j + 1 < len(nonblank) else ""
            prv = nonblank[j - 1] if j else ""
            heading = BLOCK_RE.match(nxt) or SUBBLOCK_RE.match(nxt) or SUBBLOCK_NOUNITS_RE.match(nxt) or not nxt
            if (nxt.strip().upper() in ("OR", "AND") or heading) and not COURSE_NOUNITS_RE.match(prv):
                                                                 # units with no course: a blank catalog row
                cur.setdefault("blank_rows", []).append(
                    (float(UNITS_ONLY_RE.match(l).group(1)), nxt.strip().upper() in ("OR", "AND")))
            if connector != "OR":
                slot = option = connector = None
            continue
        if s.upper() in ("OR", "AND"):
            connector = s.upper()
            continue
        if WRAPPED_TITLE_RE.search(s):                           # second line of a wrapped course title
            continue
        if l.startswith(" " * 10) and not re.search(r"\d", s) and option is not None and connector is None:
            continue                                             # wrapped title without an "(Active)" tag
        if not cur["slots"] and re.search(r"select|complete|choose|units|courses", s, re.I) and len(s) < 140:
            cur["header"] += " " + s
        elif len(s) < 45 and not s.endswith(".") and not re.search(r"\d", s):
            area = s.title() if s.isupper() else s               # area sub-heading, e.g. "Team Sports"
            cur["notes"].append(f"[area] {area}")
            slot = option = connector = None
        else:
            cur["notes"].append(s)
    return blocks, courses


def count_connectors(req_lines):
    """'or'/'and' joining courses in the catalog text: lone OR/AND lines and course lines ending in 'or'."""
    n = 0
    lines = [l for l in req_lines if l.strip() and not UNITS_ONLY_RE.match(l)]
    for i, l in enumerate(lines):
        s = l.strip().lower()
        nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if s in ("or", "and") and nxt and not re.match(r"[*#]*[A-Z]{2,5}\s+[A-Z]?\d", nxt):
            continue                                             # "or" between specialization headings
        if s in ("or", "and") or (s.endswith(" or") and re.match(r"[a-z]{2,5}\s+[a-z]?\d", s)):
            n += 1
        elif s.endswith(" or") and l.startswith("   "):            # wrapped title ending in "or"
            n += 1
    return n


def meta_records(layout_text, plain_text):
    """CurriQunet META layout (2025-26 onward): one record per program."""
    lines = clean_lines(layout_text)
    desc_units = description_units(plain_text)
    headers = find_program_headers(lines)
    records = []
    for k, (_, he, header) in enumerate(headers):
        end = headers[k + 1][0] if k + 1 < len(headers) else len(lines)
        seg = lines[he + 1:end]
        award = next((short for pat, short in AWARDS if re.search(pat, header, re.I)), "Other")
        name = re.split(r"\s+-\s+(?=ASSOCIATE|CERTIFICATE|BACHELOR|LICENSURE|Associate|Certificate)", header)[0]
        if award == "Licensure only":
            name = header.split(":")[0]
        name = name.strip().title() if name.isupper() else name.strip()

        desc, in_sum = [], False
        for l in seg:
            s = l.strip()
            if s == "Summary":
                in_sum = True
                continue
            if in_sum and re.match(r"^(Award Notes?:|Learning Outcome|Requirements$|NOTE:|Accreditation$)", s):
                break
            if in_sum and s:
                desc.append(s)

        rec = {"program": name, "award": award, "description": " ".join(desc),
               "total_lo": None, "total_hi": None, "blocks": [], "courses": {}, "missing_requirements": False}
        records.append(rec)
        r0 = next((i for i, l in enumerate(seg) if l.strip() == "Requirements"), None)
        if r0 is None:                                           # some pages omit the heading
            r0 = next((i - 1 for i, l in enumerate(seg) if BLOCK_RE.match(l)), None)
        if r0 is None:
            rec["missing_requirements"] = True
            continue
        req = []
        for l in seg[r0 + 1:]:
            tm = TOTAL_RE.search(l)
            if tm:
                rec["total_lo"], rec["total_hi"] = float(tm.group(1)), float(tm.group(2) or tm.group(1))
                break
            if l.strip() == "Courses":
                break
            req.append(l)

        rec["blocks"], rec["courses"] = parse_requirement_lines(req, desc_units)
        rec["connectors"] = count_connectors(req)
    return records


def infer_rule_a(blocks, total_lo):
    """A list with no 'select' line that adds up to more than the catalog total means
    'choose enough units from this list to reach the total' (e.g., 8 courses x 3 units, total 9)."""
    if total_lo is None:
        return
    implied = []
    for b in blocks:
        rule, minv, _ = classify(b["header"])
        if rule == "TOPUP":
            rule, minv = "UNITS", max(0.0, minv - sum(implied))
        slot_min = [min(sum(c["units"] for c in op) for op in sl) for sl in b["slots"]]
        if rule == "ALL" or (rule is None and len(b["slots"]) == 1 and len(b["slots"][0]) > 1):
            implied.append(sum(slot_min))
        elif rule == "UNITS":
            implied.append(minv)
        elif rule == "COURSES":
            implied.append(sum(sorted(sum(c["units"] for c in op) for sl in b["slots"] for op in sl)[:int(minv)]))
        else:
            implied.append(b["lo"])
    if sum(implied) <= total_lo + 0.05:
        return
    plain = [i for i, b in enumerate(blocks)
             if classify(b["header"])[0] == "ALL" and not b["sub"] and len(b["slots"]) > 1
             and all(len(sl) == 1 for sl in b["slots"])]
    if len(plain) != 1:
        return
    i = plain[0]
    need = total_lo - (sum(implied) - implied[i])
    if 0 < need < implied[i]:
        blocks[i]["header"] = f"Complete {need:g} units from the following: [inferred from catalog total]"


def records_to_tables(records, year):
    """Shared by both layouts: rules, checksum, flags, and the flat tables. IDs carry the catalog year."""
    yy = year[2:4] + "-" + year[-2:] if year else ""
    prog_rows, block_rows, opt_rows, flags, courses = [], [], [], [], {}
    for k, rec in enumerate(records, start=1):
        pid = f"{yy}-{k}" if yy else k
        name, award, blocks = rec["program"], rec["award"], rec["blocks"]
        total_lo, total_hi = rec["total_lo"], rec["total_hi"]
        if rec["missing_requirements"]:                          # umbrella heading; its options are parsed separately
            continue
        for code, info in rec["courses"].items():
            courses.setdefault(code, info)
        prog_flags, implied_min, subj_counts = [], 0.0, {}
        infer_rule_a(blocks, total_lo)
        for bi, b in enumerate(blocks, start=1):
            bid = f"{pid}-{bi}"
            rule, minv, _ = classify(b["header"])
            reasons = []
            hl = re.sub(r"\s+", " ", b["header"].lower())
            am = re.search(r"from (?:any )?(\w+) (?:of )?(?:the following )?areas", hl)
            areas_required = num(am.group(1)) if am else None
            rng = re.search(r"(\d+(?:\.\d+)?)\s*(?:-|–|to)\s*(\d+(?:\.\d+)?)\)?\s*units", hl)
            hi = b["hi"] or (float(rng.group(2)) if rng else 0.0)
            if areas_required:
                rule, minv = "COURSES", int(areas_required)
            no_double = bool(re.search(r"not (already )?(selected|completed)", hl))
            slot_min = [min(sum(c["units"] for c in op) for op in sl) for sl in b["slots"]]
            if rule == "TOPUP":                                  # units still needed to reach the stated total
                rule, minv = "UNITS", max(0.0, minv - implied_min)
            if rule is None and len(b["slots"]) == 1 and len(b["slots"][0]) > 1:
                rule = "ALL"                                     # e.g. "Select all BUSE 90 courses ... OR BUSE 277C"
            if rule is None:
                rule, minv = "UNITS", b["lo"]
                reasons.append(f"rule not recognized: \"{b['header'][:80]}\"")
            if rule == "ALL":
                minv = len(b["slots"])
                implied = sum(slot_min)
            elif rule == "UNITS":
                implied = minv
            else:
                opt_units = sorted(sum(c["units"] for c in op) for sl in b["slots"] for op in sl)
                implied = sum(opt_units[:int(minv)])
            implied_min += implied
            if not b["slots"]:
                reasons.append("group has no courses")
            if (rule == "ALL" and b["lo"] and abs(implied - b["lo"]) > 0.05
                    and not (bi < len(blocks) and blocks[bi]["sub"])):
                reasons.append(f"catalog section says {b['lo']:g} units but its courses add up to {implied:g}")
            if no_double:
                reasons.append("not enforced yet: a course can't count in two groups (\"not selected above\")")
            if areas_required:
                reasons.append(f"not enforced yet: courses must come from {int(areas_required)} different areas")
            review = bool(reasons)
            prog_flags += [f"Group {bi}: {r}" for r in reasons]
            for sl in b["slots"]:
                for op in sl:
                    for c in op:
                        sj = c["code"].split()[0]
                        subj_counts[sj] = subj_counts.get(sj, 0) + 1
            block_rows.append({"block_id": bid, "program_id": pid, "block_order": bi, "header": b["header"],
                               "rule_type": rule, "min_required": minv, "catalog_units_max": hi,
                               "areas_required": areas_required, "no_double_count": no_double,
                               "within_required_section": rule == "ALL" or b["sub"], "needs_review": review})
            for si, sl in enumerate(b["slots"], start=1):
                for oi, op in enumerate(sl, start=1):
                    for c in op:
                        opt_rows.append({"block_id": bid, "slot_id": f"{bid}-{si}", "option_id": f"{bid}-{si}-{oi}",
                                         "course": c["code"], "units": c["units"], "units_max": c["units_max"],
                                         "area": c["area"]})

        built = sum(len(sl) - 1 + sum(len(op) - 1 for op in sl) for b in blocks for sl in b["slots"])
        blank = [r for b in blocks for r in b.get("blank_rows", [])]
        built += sum(1 for _, joined in blank if joined)         # an OR/AND to a blank row is the catalog's gap
        if rec.get("connectors", 0) != built:
            prog_flags.append(f"OR/AND check: catalog text has {rec.get('connectors', 0)} connectors, data has {built}")
        checksum_ok = total_lo is not None and abs(implied_min - total_lo) <= 0.05
        upper_only = False
        if award == "BS" and total_lo is not None and not checksum_ok:   # BS total can count upper division only
            upper = sum(c["units"] for b in blocks for sl in b["slots"] for op in sl[:1] for c in op
                        if int(re.sub(r"\D", "", c["code"].split()[1]) or 0) >= 300)
            if abs(upper - total_lo) <= 0.05:
                checksum_ok = upper_only = True
        if total_lo is None:
            prog_flags.append("No Total Units line found")
        elif not checksum_ok:
            prog_flags.append(f"Parsed minimum {implied_min:g} units vs catalog Total Units {total_lo:g}")
        if not blocks:
            prog_flags.append("No requirement blocks parsed")
        notes = []
        if any("prerequisite" in b["header"].lower() for b in blocks):
            notes.append("Selective admission: students must complete program prerequisites and be accepted into the "
                         "program. See a counselor.")
        if re.search(r"general education|igetc|getc", name, re.I):
            notes.append("General education pattern, not a major: almost any GE course counts toward it, so it can look "
                         "close for most students. See a counselor.")
            prog_flags.append("Flag: general education pattern")
        if re.search(r"apprentice", rec["description"] + " " + name + " " + rec.get("department", ""), re.I):
            notes.append("Apprenticeship: open only to indentured apprentices in the trade program. See a counselor.")
            prog_flags.append("Flag: apprenticeship program")
        if any(re.search(r"(?:from|select) one of the following speciali", b["header"], re.I) for b in blocks):
            notes.append("Specialization units must all come from ONE specialization; the dashboard does not check "
                         "this yet. Verify with a counselor.")
            prog_flags.append("Flag: single-specialization rule not enforced")
        mine = [r for r in block_rows if r["program_id"] == pid]
        if any(r["no_double_count"] for r in mine):
            notes.append("A course can count toward only one group in this program; the dashboard does not check "
                         "this yet, so progress may look better than it is. Verify with a counselor.")
        if any(r["areas_required"] for r in mine):
            notes.append("Some courses must come from different areas; the dashboard does not check this yet. "
                         "Verify with a counselor.")
        if upper_only:
            notes.append(f"The catalog's {total_lo:g} total units count upper-division (300+) courses only; the "
                         f"lower-division major courses listed are also required ({implied_min - total_lo:g} units).")
        if blank:
            notes.append(("The catalog shows a blank course row" if len(blank) == 1 else
                          f"The catalog shows {len(blank)} blank course rows") + " ("
                         + "; ".join(f"{u:g} units" + (", listed as an OR alternative" if j else "") for u, j in blank)
                         + ") with no course named. Verify with a counselor.")
            prog_flags.append(f"Flag: {len(blank)} blank course row(s) in the catalog")
        if total_lo is not None and not checksum_ok:
            notes.append(f"The catalog lists {total_lo:g} total units, but the listed requirements add up to a "
                         f"minimum of {implied_min:g}. Verify with a counselor.")
        note = " ".join(notes)
        flags += [{"catalog_year": year, "program_id": pid, "program": name, "award": award, "issue": fl}
                  for fl in prog_flags]
        prog_rows.append({"catalog_year": year, "program_id": pid, "program": name, "award": award,
                          "department": rec.get("department", ""), "program_note": note,
                          "main_subject": max(subj_counts, key=subj_counts.get) if subj_counts else "",
                          "total_units_min": total_lo, "total_units_max": total_hi, "parsed_units_min": implied_min,
                          "checksum_ok": checksum_ok, "needs_review": bool(prog_flags),
                          "description": rec["description"]})

    course_df = pd.DataFrame([{"course": k, "subject": v["subject"], "title": v["title"],
                               "units_min": v["units"], "units_max": v["units_max"]}
                              for k, v in sorted(courses.items())])
    return (pd.DataFrame(prog_rows), pd.DataFrame(block_rows), pd.DataFrame(opt_rows),
            course_df, pd.DataFrame(flags))


# ---------------------------------------------------------------- classic two-column layout (2020-21 to 2024-25)
CLASSIC_AWARD_RE = re.compile(
    r"^(Associate (?:of|in) (?:Arts|Science)(?: for Transfer)?(?: Degree)?|Certificate of (?:Achievement|Performance)"
    r"|Bachelor of Science(?: Degree)?)\s*\**\s*(?::\s*(.*)|$)")   # "Certificate of Performance*:"
CLASSIC_COURSE_RE = re.compile(
    r"^([A-Z]{2,5})\s+([A-Z]?[0-9]{1,4}[A-Z]{0,3})\s+(.+?)\s{2,}([0-9]+(?:\.[0-9]+)?)(?:\s*[–-]\s*([0-9]+(?:\.[0-9]+)?))?\s*$")
# Long titles can leave a single space before the units: "ANTH 103 Introduction to Cultural Anthropology 3"
CLASSIC_COURSE_TIGHT_RE = re.compile(
    r"^([A-Z]{2,5})\s+([A-Z]?[0-9]{1,4}[A-Z]{0,3})\s+(.+?[A-Za-z).,])\s+([0-9]{1,2}(?:\.[0-9]+)?)(?:\s*[–-]\s*([0-9]{1,2}(?:\.[0-9]+)?))?\s*$")
# Title ending in a number, then units: "CISC 193 Microsoft C# Software Engineering 1 4"
CLASSIC_COURSE_NUMTITLE_RE = re.compile(
    r"^([A-Z]{2,5})\s+([A-Z]?[0-9]{1,4}[A-Z]{0,3})\s+(.+?\s[0-9IVX]{1,3})\s+([0-9]{1,2}(?:\.[0-9]+)?)()\s*$")
CLASSIC_COURSE_NOUNITS_RE = re.compile(r"^([A-Z]{2,5})\s+([A-Z]?[0-9]{1,4}[A-Z]{0,3})\s+(\S.*?)\s*$")
CLASSIC_CONT_RE = re.compile(r"^\s{3,}(\S.*?)?\s*([0-9]+(?:\.[0-9]+)?)(?:\s*[–-]\s*([0-9]+(?:\.[0-9]+)?))?\s*$")
CLASSIC_TOTAL_RE = re.compile(r"Total Units\s*=?\s*([0-9.]+)(?:\s*[–-]\s*([0-9.]+))?")
CLASSIC_REQ_START_RE = re.compile(
    r"^(?:Courses(?: Required(?: for(?: the)? Major)?)?|\**Program Prerequisites)\s*:?\s*(?:Units)?$", re.I)
FOOTER_RE = re.compile(r"San Diego City College\s*•")
CLASSIC_AWARD_SHORT = {"Associate in Arts for Transfer": "AA-T", "Associate in Science for Transfer": "AS-T",
                       "Associate of Arts": "AA", "Associate of Science": "AS",
                       "Certificate of Achievement": "Certificate of Achievement",
                       "Certificate of Performance": "Certificate of Performance", "Bachelor of Science": "BS"}


def column_gutters(pdf):
    """Per page: (width, height, x of the emptiest vertical strip near the middle) from word positions."""
    html = subprocess.run(["pdftotext", "-bbox", str(pdf), "-"], capture_output=True).stdout.decode("utf-8", "replace")
    pages = []
    for pg in re.finditer(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', html, re.S):
        w, h = float(pg.group(1)), float(pg.group(2))
        cover = [0] * int(w + 1)
        for x0, x1 in re.findall(r'<word xMin="([\d.]+)" yMin="[\d.]+" xMax="([\d.]+)"', pg.group(3)):
            for x in range(int(float(x0)), min(int(float(x1)) + 1, len(cover))):
                cover[x] += 1
        lo, hi = int(w * 0.38), int(w * 0.62)
        best = min(range(lo, hi), key=lambda x: (sum(cover[x - 2:x + 3]), abs(x - w / 2)))
        pages.append((w, h, best))
    return pages


def classic_lines(pdf):
    """Read each page as two columns, left then right, with each column's own margin removed."""
    out = []
    for p, (w, h, gutter) in enumerate(column_gutters(pdf), start=1):
        for x0, cw in ((0, gutter), (gutter, w - gutter)):
            col = subprocess.run(["pdftotext", "-layout", "-f", str(p), "-l", str(p), "-x", str(int(x0)), "-y", "0",
                                  "-W", str(int(cw) + 1), "-H", str(int(h)), str(pdf), "-"],
                                 capture_output=True).stdout.decode("utf-8", "replace").replace("\f", "")
            rows = [l.rstrip() for l in col.splitlines()]
            leads = sorted(len(l) - len(l.lstrip()) for l in rows if len(l.strip()) >= 25)
            indent = leads[len(leads) // 4] if leads else 0          # typical left edge; ignores stray margin labels
            for l in rows:
                lead = len(l) - len(l.lstrip())
                s = l.strip()
                if s.lower() in ("or", "and"):
                    out.append(" " * 10 + s)
                    continue
                if (FOOTER_RE.search(s) or s == "Academic Programs" or re.fullmatch(r"\d{1,3}", s)
                        or (lead < indent - 3 and len(s) < 60 and not re.search(r"\d", s))   # margin label
                        or (lead - indent >= 40 and not re.search(r"\d", s))):
                    continue
                out.append(l[min(lead, indent):])
    return out


def classic_records(pdf):
    lines = classic_lines(pdf)
    for i in range(len(lines)):                                  # "Associate in Arts in X" ... "for Transfer Degree:"
        a = lines[i].strip()
        if not re.match(r"Associate in (Arts|Science) in\b", a) or re.search(r"for Transfer(?: Degree)?\s*:?$", a):
            continue
        parts, k = [a], i + 1
        while k < len(lines) and k <= i + 3 and lines[k].strip() and len(lines[k].strip()) < 60:
            parts.append(lines[k].strip())
            if re.search(r"(for Transfer(?: Degree)?|Degree)\s*:?$", parts[-1]):
                joined = " ".join(parts)
                if re.search(r"for Transfer(?: Degree)?\s*:?$", joined):
                    lines[i] = joined.rstrip(": ") + ":"
                    for z in range(i + 1, k + 1):
                        lines[z] = ""
                break
            k += 1
    starts = []
    for i, l in enumerate(lines):
        t = re.match(r"^Associate in (Arts|Science) in (.+?) for Transfer(?: Degree)?\s*:\s*$", l.strip())
        if t:
            award = "AA-T" if t.group(1) == "Arts" else "AS-T"
            starts.append((i, i + 1, t.group(2).strip(), award))
            continue
        m = CLASSIC_AWARD_RE.match(l.strip())
        if not m or len(l) - len(l.lstrip()) > 25:
            continue
        if m.group(2) is None and any(re.search(r"\s\d+(?:\.\d+)?(?:[–-]\d+(?:\.\d+)?)?\*?$", x)
                                      for x in lines[i + 1:i + 4]):
            continue
        name_lines = [m.group(2).strip()] if (m.group(2) or "").strip() else []
        j = i + 1
        while not name_lines or (len(name_lines) < 4 and not name_lines[-1].endswith("*")):
            if j < len(lines) and not lines[j].strip() and name_lines:   # column break inside the name
                k2 = next((z for z in range(j, min(j + 6, len(lines))) if lines[z].strip()), None)
                if (k2 is not None and k2 + 1 < len(lines) and len(lines[k2].split()) <= 3
                        and not re.search(r"[\d.:]", lines[k2])
                        and re.match(r"(Description|Courses Required|Program Prerequisites)", lines[k2 + 1].strip())):
                    j = k2
            if j >= len(lines) or not lines[j].strip():
                break
            nxt = lines[j].strip()
            if name_lines and (len(nxt) > 45 or re.match(r"(The|This|Students|Program|Upon|Note|Award|Career|Courses|Description|"
                                                         r"Prepares|Prerequisite|Requirements|Associate|Certificate|Completion)\b", nxt)
                               or re.search(r"\b(will|this|is|are)\b", nxt)       # a sentence, not a name
                               or (len(nxt) >= 20 and nxt.lower().startswith(name_lines[0][:12].lower()))
                               or re.search(r"\s{2,}\d", nxt) or nxt.endswith(".")):
                break
            name_lines.append(nxt)
            j += 1
        name = " ".join(name_lines)
        if re.fullmatch(r"Requirements:?", name):                # "Certificate of Achievement Requirements:" section
            continue
        if (not name or re.search(r"\s{2,}\d+(?:\.\d+)?(?:[–-]\d+(?:\.\d+)?)?\*?$", name) or re.search(r"\s{3,}\d", name)
                or CLASSIC_AWARD_RE.search(name)):                # department summary table row, not a program
            continue
        award = next(v for k, v in CLASSIC_AWARD_SHORT.items() if m.group(1).startswith(k))
        starts.append((i, j, re.sub(r"\*+$", "", name).strip(), award))
    context = {i: " ".join(x.strip() for x in lines[max(0, i - 120):i]) for i, *_ in starts}
    dept_at = []                                                 # department name sits just above "Award Type  Units"
    for t, l in enumerate(lines):
        if re.match(r"\s*Award Type\s+Units", l):
            title = []
            for x in reversed(lines[max(0, t - 6):t]):
                x = x.strip()
                if not x:
                    continue
                if len(x) > 45 or x.endswith(".") or re.search(r"\d", x):
                    break
                title.insert(0, x)
            if title:
                dept_at.append((t, " ".join(title)))
    def dept_of(i):
        prior = [d for t, d in dept_at if t < i]
        return prior[-1] if prior else ""

    records = []
    for k, (i, j, name, award) in enumerate(starts):
        d = dept_of(i)
        rest = name[len(d) + 1:] if d and name.startswith(d + " ") else ""
        if rest and any(w[:4].lower() in rest.lower() for w in re.findall(r"[A-Za-z]{4,}", d)):
            name = rest                                           # "Nursing Education | Registered Nurse: Generic"
        end = starts[k + 1][0] if k + 1 < len(starts) else len(lines)
        seg = lines[j:end]
        r0 = next((n for n, l in enumerate(seg) if CLASSIC_REQ_START_RE.match(l.strip())), None)
        desc = []
        for l in seg[:r0 if r0 is not None else len(seg)]:
            s = l.strip()
            if not s or re.match(r"(Program Learning Outcomes|Award Notes?|Note|Career Options|Program Goals)", s):
                if desc:
                    break
                continue
            desc.append(s)
        rec = {"program": name, "award": award, "description": " ".join(desc), "total_lo": None, "total_hi": None,
               "context": context.get(i, ""), "department": dept_of(i),
               "blocks": [], "courses": {}, "missing_requirements": r0 is None}
        records.append(rec)
        if r0 is None:
            continue
        req = []
        for l in seg[r0 + 1:]:
            tm = CLASSIC_TOTAL_RE.search(l)
            if tm:
                rec["total_lo"], rec["total_hi"] = float(tm.group(1)), float(tm.group(2) or tm.group(1))
                break
            req.append(l)
        cut = next((n for n, l in enumerate(req)               # suggestions, not requirements
                    if re.match(r"(The following .*)?recommended electives", l.strip(), re.I)), None)
        if cut is not None:
            req = req[:cut]
        header = seg[r0].strip()
        rec["blocks"], rec["courses"] = parse_classic_requirements(req, header)
        rec["connectors"] = count_connectors(req)
    return records


def parse_classic_requirements(req, first_header):
    blocks, slot, option, connector, area, courses, pending, held = [], None, None, None, "", {}, None, []
    cur = {"header": first_header, "lo": 0.0, "hi": 0.0, "slots": [], "notes": [], "sub": False}
    blocks.append(cur)

    def place(code, subj, title, u1, u2):
        nonlocal slot, option, connector, held
        if held:                                                 # "X ... or" (no units) shares units with its alternative
            waiting, held = held, []
            for h in waiting:
                place(f"{h[0]} {h[1]}", h[0], h[2], u1, u2)
                connector = "OR"
        c = {"code": code, "units": float(u1), "units_max": float(u2 or u1), "area": area}
        courses.setdefault(code, {"subject": subj, "title": title.strip(), "units": c["units"],
                                  "units_max": c["units_max"]})
        if connector == "AND" and option is not None:
            option.append(c)
        elif connector == "OR" and slot is not None:
            option = [c]
            slot.append(option)
        else:
            option = [c]
            slot = [option]
            cur["slots"].append(slot)
        connector = None

    for l in req:
        l = re.sub(r"\s*[*#†‡]+(?=\s|$)", "", l.rstrip())       # footnote marks: "Mechanics *#  5"
        l = re.sub(r"^(\s*)[*#†‡]+(?=[A-Z]{2,5}\s+[A-Z]?\d)", r"\1", l)   # leading marks: "**ENGL 215 ..."
        l = re.sub(r"^(\s*[A-Z]{2,5}\s+[A-Z]?\d{1,4}[A-Z]?)(?=[A-Z][a-z])", r"\1 ", l)   # "ELDT 123LDigital"
        s = l.strip()
        if not s:
            continue
        if pending and s.lower() == "or":                        # "X ... <wrapped title>" then a lone "or"
            held.append(pending)
            pending = None
            continue
        if pending:                                              # title wrapped onto this line
            cm = CLASSIC_CONT_RE.match(l)
            if cm:
                subj, cnum, title = pending
                place(f"{subj} {cnum}", subj, title + " " + (cm.group(1) or ""), cm.group(2), cm.group(3))
                pending = None
                continue
            if l.startswith("   ") and not CLASSIC_COURSE_NOUNITS_RE.match(s):
                if s.lower().endswith(" or"):
                    held.append((pending[0], pending[1], pending[2] + " " + s[:-3]))
                    pending = None
                else:
                    pending = (pending[0], pending[1], pending[2] + " " + s)
                continue
            cur["notes"].append(f"{pending[0]} {pending[1]}: units not found")
            pending = None
        m = CLASSIC_COURSE_RE.match(s) or CLASSIC_COURSE_TIGHT_RE.match(s) or CLASSIC_COURSE_NUMTITLE_RE.match(s)
        if m:
            subj, cnum, title, u1, u2 = m.groups()
            place(f"{subj} {cnum}", subj, title, u1, u2)
            continue
        m = CLASSIC_COURSE_NOUNITS_RE.match(s)
        if m and s.lower().endswith(" or"):                      # "ASTR 109 Practice in Observing or"
            subj, cnum, title = m.groups()
            held.append((subj, cnum, title[:-3].strip()))
            continue
        if m:
            pending = m.groups()
            continue
        if s.lower() in ("or", "and"):
            connector = s.upper()
            continue
        if s.lower().endswith(" or") and cur["slots"]:          # wrapped title ending in "or"
            connector = "OR"
            continue
        if (re.match(r"(core|additional\b.*\brequired|.*\bcourses? required)\s*:?$", s, re.I)
                or re.match(r"(courses required for|\**program prerequisites)\b", s, re.I)):
            cur = {"header": s, "lo": 0.0, "hi": 0.0, "slots": [], "notes": [], "sub": False}
            blocks.append(cur)
            slot = option = connector = None
            area = ""
            continue
        if re.match(r"(select|complete|choose|students (should|must) complete)\b", s, re.I) or \
                re.match(r"if needed to (total|reach)\b.*\b(select|complete|choose)\b", s, re.I) or \
                re.match(r"\d+(\.\d+)?(\s*(to|-|–)\s*\d+(\.\d+)?)?\s+(semester\s+)?units?\b.*(from|following)", s, re.I):
            cur = {"header": s, "lo": 0.0, "hi": 0.0, "slots": [], "notes": [], "sub": True}
            blocks.append(cur)
            slot = option = connector = None
            area = ""
            continue
        if (not cur["slots"] and cur["sub"] and len(s) < 140 and not re.match(r"\(", s)
                and not cur["header"].rstrip().endswith(":")):
            cur["header"] += " " + s                             # wrapped "Select ... from the" / "following:"
            continue
        if len(s) < 45 and not s.endswith(".") and not re.search(r"\d", s) and not s.startswith("("):
            area = s
            cur["notes"].append(f"[area] {area}")
            slot = option = connector = None
            continue
        cur["notes"].append(s)
    if pending:
        cur["notes"].append(f"{pending[0]} {pending[1]}: units not found")
    for h in held:
        cur["notes"].append(f"{h[0]} {h[1]}: 'or' with no alternative after it")
    return [b for b in blocks if b["slots"] or b["sub"]], courses


# ---------------------------------------------------------------- output tables
def build_program_courses(programs, blocks, options, courses):
    prog = programs.set_index("program_id")
    titles = courses.set_index("course").title.to_dict()
    out = []

    def add(bl, rtype, gid, gname, needed, unit, n_opts, opt):
        seq = " + ".join(opt.course) if len(opt) > 1 else ""
        p = prog.loc[bl.program_id]
        if rtype == "Required":
            rule = "Required"
        elif bl.areas_required == bl.areas_required and bl.areas_required:
            rule = f"Complete {int(needed)} courses, each from a different area ({n_opts} options)"
        elif unit == "units" and bl.catalog_units_max > needed:
            rule = f"Complete {needed:g} to {bl.catalog_units_max:g} units from {n_opts} options"
        elif unit == "units":
            rule = f"Complete {needed:g} units from {n_opts} options"
        else:
            rule = f"Complete {int(needed)} of {n_opts} options"
        for _, x in opt.iterrows():
            out.append({"catalog_year": p.catalog_year, "program_id": bl.program_id, "program": p.program, "award": p.award,
                        "requirement_type": rtype, "group_id": gid, "option_id": x.option_id, "group_name": gname, "group_rule": rule,
                        "needed_from_group": needed, "needed_unit": unit, "options_in_group": n_opts,
                        "course": x.course, "course_title": titles.get(x.course, ""),
                        "units_min": x.units, "units_max": x.units_max,
                        "units_display": f"{x.units:.1f}" if x.units == x.units_max
                        else f"{x.units:.1f}-{x.units_max:.1f}",
                        "taken_together_with": seq, "area": x.area,
                        "areas_required": bl.areas_required, "no_double_count": bool(bl.no_double_count),
                        "catalog_wording": bl.header,
                        "needs_review": bool(bl.needs_review) or not bool(p.checksum_ok)})

    for pid, pb in blocks.groupby("program_id", sort=False):
        gnum = 0
        for _, bl in pb.sort_values("block_order").iterrows():
            bo = options[options.block_id == bl.block_id]
            if bl.rule_type == "ALL":
                for _, sl in bo.groupby("slot_id", sort=False):
                    opts = [g for _, g in sl.groupby("option_id", sort=False)]
                    if len(opts) == 1:
                        add(bl, "Required", f"{pid}-REQ", "Required courses", None, "all", None, opts[0])
                    else:
                        gnum += 1
                        unit = "sequence" if any(len(g) > 1 for g in opts) else "course"
                        for g in opts:
                            add(bl, "Required Selective", f"{pid}-G{gnum}", f"Group {gnum}", 1, unit, len(opts), g)
            else:
                gnum += 1
                opts = [g for _, g in bo.groupby("option_id", sort=False)]
                rtype = "Required Selective" if bl.within_required_section else "Elective"
                unit = "units" if bl.rule_type == "UNITS" else "courses"
                for g in opts:
                    add(bl, rtype, f"{pid}-G{gnum}", f"Group {gnum}", bl.min_required, unit, len(opts), g)
    return pd.DataFrame(out)


def build_review(programs, blocks, options):
    prog = programs.set_index("program_id")
    rows = []
    for _, bl in blocks.iterrows():
        p = prog.loc[bl.program_id]
        rule = {"ALL": "All required", "UNITS": f"Choose ≥ {bl.min_required:g} units",
                "COURSES": f"Choose {int(bl.min_required)} course(s)/option(s)"}[bl.rule_type]
        parts = []
        for _, sl in options[options.block_id == bl.block_id].groupby("slot_id", sort=False):
            opts = [" + ".join(g.course) for _, g in sl.groupby("option_id", sort=False)]
            parts.append(" OR ".join(opts) if len(opts) == 1 or bl.rule_type != "ALL" else f"({' OR '.join(opts)})")
        rows.append({"program_id": bl.program_id, "program": p.program, "award": p.award, "block": bl.block_order,
                     "catalog_header": bl.header, "parsed_rule": rule, "courses": "; ".join(parts),
                     "checksum_ok": p.checksum_ok, "needs_review": bl.needs_review})
    return pd.DataFrame(rows)


def build_readme(source, programs, courses, program_courses):
    items = {
        "Source": f"San Diego City College catalogs, parsed {dt.date.today():%Y-%m-%d}. {source}.",
        "Catalog year": "catalog_year marks which catalog each row comes from. IDs include the year (e.g., 23-24-63), "
                        "so filter the dashboard to one catalog year at a time.",
        "Coverage": f"{len(programs)} programs, {len(courses)} distinct courses, {len(program_courses)} program-course "
                    "rows. Major requirements only: GE, the 60-unit degree minimum, grade rules and catalog rights "
                    "are NOT included.",
        "Checksum": "Parsed minimum units were compared to each program's catalog \"Total Units\" line. "
                    f"{int(programs.checksum_ok.sum())} of {len(programs)} match exactly.",
        "Needs review": f"{int(programs.needs_review.sum())} programs have at least one flag (ReviewFlags tab). "
                        "Their rows are marked needs_review = TRUE in ProgramCourses.",
        "ProgramCourses": "One row per program × course. This is the bridge between Programs and Courses. option_id groups "
                          "courses that count together as one option (a sequence); most options are a single course.",
        "Requirement types": "Required = must take. Required Selective = a choice inside the catalog's required "
                             "courses (e.g., ARTF 104 OR ARTH C1200). Elective = a separate \"select N from the "
                             "following\" list. For both selective types, needed_from_group says how much of the "
                             "group completes it.",
        "Groups": "needed_from_group + needed_unit say what completes a group: e.g., 7 units, or 2 courses. A program "
                  "can have several groups; group_id keeps them apart. Required courses share one group per program.",
        "Variable units": "Variable-unit courses (e.g., CHIL 270 Work Experience, 1.0-4.0) have units_min < units_max. "
                          "The course selector should count units_min (conservative); with transcript data, use "
                          "units actually earned.",
        "Sequences": "taken_together_with lists courses that count only as a set (e.g., PHYS 180A + PHYS 181A + "
                     "PHYS 180B + PHYS 181B). Completing part of a set does not count.",
        "Areas": "area is the catalog sub-heading a course sits under (e.g., Aerobic, Team Sports, Infant/Toddler). "
                 "areas_required > 0 means the group must be completed with courses from that many different areas.",
        "No double count": "no_double_count = TRUE marks groups worded \"not selected/completed above\": a course already "
                           "used for an earlier group of the same program cannot also count here.",
        "Not modeled yet": "Prerequisites, and whether specialization groups (e.g., Child Development: Master Teacher) "
                           "must be completed within one area, are not enforced.",
        "Program note": "program_note (Programs tab) carries cautions that apply to the whole program. Programs with a "
                        "\"Program Prerequisites\" section (e.g., Nursing) are selective admission; their prerequisites "
                        "are included in the requirements so progress can be tracked as if the student were in the program.",
        "Department": "The catalog does not tag programs with a department. main_subject (Programs tab) is the most "
                      "frequent course prefix and should be replaced with an official mapping.",
    }
    return pd.DataFrame({"Item": list(items), "Detail": list(items.values())})


def write_workbook(path, sheets):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, df in sheets:
            df.to_excel(w, sheet_name=name, index=False)
    wb = load_workbook(path)
    widths = {"README": {"A": 18, "B": 130},
              "ProgramCourses": {"B": 40, "C": 22, "D": 24, "E": 10, "F": 14, "G": 14, "H": 34, "L": 12, "M": 42,
                                 "Q": 40, "R": 20, "U": 45},
              "Programs": {"B": 45, "C": 24, "J": 80}, "Courses": {"C": 50},
              "Review": {"B": 38, "C": 22, "E": 45, "F": 26, "G": 90}, "ReviewFlags": {"B": 45, "C": 24, "D": 90}}
    hdr_fill, flag_fill = PatternFill("solid", fgColor="1F3864"), PatternFill("solid", fgColor="FCE4D6")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Arial", size=10)
        for cell in ws[1]:
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.fill = hdr_fill
        for col, wd in widths.get(ws.title, {}).items():
            ws.column_dimensions[col].width = wd
        ws.freeze_panes = "A2"
        if ws.title == "README":
            for cell in ws["B"]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            continue
        table = Table(displayName=ws.title, ref=ws.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
        ws.add_table(table)
        hdr = [c.value for c in ws[1]]
        if "needs_review" in hdr:
            ci = hdr.index("needs_review")
            ck = hdr.index("checksum_ok") if "checksum_ok" in hdr else None
            for row in ws.iter_rows(min_row=2):
                if row[ci].value is True or (ck is not None and row[ck].value is False):
                    for cell in row:
                        cell.fill = flag_fill
    wb.save(path)


YEAR_RE = re.compile(r"(20\d\d)\D{0,3}(20\d\d)")


def catalog_year(path):
    m = YEAR_RE.search(path.name)
    if not m:
        raise SystemExit(f"Can't find a catalog year like 2023-2024 in the file name: {path.name}")
    return f"{m.group(1)}-{m.group(2)}"


def parse_pdf(pdf):
    year = catalog_year(pdf)
    layout = pdf_to_text(pdf, layout=True)
    if re.search(r":\s*CITY\s*$", layout, re.M | re.I):
        fmt, records = "META", meta_records(layout, pdf_to_text(pdf, layout=False))
    else:
        fmt, records = "classic", classic_records(pdf)
    return year, fmt, records_to_tables(records, year)


def collect_pdfs(inputs):
    """Folders, wildcards and file names in; one PDF per catalog year out.
    In a folder or wildcard, PDFs without a year in the name are skipped; a named file must have one."""
    found = []
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            hits = [f for f in sorted(p.glob("*.pdf")) if YEAR_RE.search(f.name)]
            if not hits:
                raise SystemExit(f"No PDFs with a catalog year in the name were found in {p}")
            found += hits
        elif any(ch in item for ch in "*?["):
            found += [f for f in sorted(Path().glob(item)) if f.suffix.lower() == ".pdf" and YEAR_RE.search(f.name)]
        elif p.exists():
            catalog_year(p)
            found.append(p)
        else:
            raise SystemExit(f"Not found: {item}")
    years = {}
    for f in found:
        y = catalog_year(f)
        if y in years and years[y] != f:
            raise SystemExit(f"Two PDFs for {y}: {years[y].name} and {f.name}. Keep one per catalog year.")
        years[y] = f
    if not years:
        raise SystemExit("No catalog PDFs to parse.")
    return [years[y] for y in sorted(years)]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+",
                    help="catalog PDFs, a folder of them, or a wildcard; the year must be in each file name")
    ap.add_argument("-o", "--output", type=Path, required=True)
    ap.add_argument("--no-description", action="store_true",
                    help="leave out the catalog's program summaries (verbatim catalog text)")
    a = ap.parse_args()

    parts, summary = [], []
    for pdf in collect_pdfs(a.pdfs):
        year, fmt, tables = parse_pdf(pdf)
        p = tables[0]
        summary.append(f"{year} ({fmt}): {len(p)} programs, {int(p.checksum_ok.sum())} match Total Units, "
                       f"{int(p.needs_review.sum())} flagged")
        print(summary[-1])
        parts.append(tables)

    programs, blocks, options, courses, flags = (pd.concat([t[i] for t in parts], ignore_index=True) for i in range(5))
    if a.no_description:
        programs = programs.drop(columns="description")
    courses = courses.drop_duplicates("course", keep="last").sort_values("course").reset_index(drop=True)
    program_courses = build_program_courses(programs, blocks, options, courses)
    review = build_review(programs, blocks, options)
    readme = build_readme("; ".join(summary), programs, courses, program_courses)

    write_workbook(a.output, [("README", readme), ("ProgramCourses", program_courses),
                              ("Programs", programs), ("Courses", courses), ("Review", review),
                              ("ReviewFlags", flags)])
    print(f"-> {a.output}")


if __name__ == "__main__":
    main()
