# -*- coding: utf-8 -*-
"""
unified_report.py — run the engine on one or more monthly גולמי files and write
a polished, workable unified workbook:

  לוח בקרה     — KPI tiles + per-file summary with conditional formatting
  שגויים לבדיקה — the work queue: only invalid slips, sorted by |gap|, filterable
  פר עובד      — every employee from every file as a filterable Excel table
  פילוח משרדים — ministry × validity aggregated across all files

Usage:
    python tools/unified_report.py file1.xlsx [file2.xlsx ...] --out unified.xlsx
"""

import argparse
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

import main as engine

# ---- palette (light surface; status colors paired with a text label, never alone)
NAVY = "FF1E3A5F"          # brand header
NAVY_TEXT = "FFFFFFFF"
TILE_BG = "FFF4F5F7"
GOOD_TXT, GOOD_BG = "FF0B7A0B", "FFE7F4E7"
BAD_TXT, BAD_BG = "FFA82626", "FFFBE9E9"
WARN_TXT, WARN_BG = "FF8A5A00", "FFFEF3D8"
MUTED = "FF6B7280"
BAR_BLUE = "FF2A78D6"
BORDER = Side(style="thin", color="FFD9DDE3")
THIN_BOX = Border(left=BORDER, right=BORDER, top=BORDER, bottom=BORDER)

MONEY = "#,##0.00"
INT = "#,##0"
PCT1 = '0.00"%"'

STATUS_HE = {"valid": "תקין", "invalid": "שגוי",
             "no_base": "ללא שכר בסיס פעיל", "multi_period": "רטרו / רב-תקופתי"}


def _header_row(ws, row, labels, widths=None):
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.font = Font(bold=True, color=NAVY_TEXT, size=11)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BOX
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 24


def _kpi(ws, row, col, span, title, value, color=None, fmt=None):
    """A merged stat tile: muted title above a large value."""
    c0, c1 = get_column_letter(col), get_column_letter(col + span - 1)
    ws.merge_cells(f"{c0}{row}:{c1}{row}")
    ws.merge_cells(f"{c0}{row + 1}:{c1}{row + 1}")
    t = ws.cell(row=row, column=col, value=title)
    t.font = Font(size=10, color=MUTED)
    t.alignment = Alignment(horizontal="center", vertical="bottom")
    v = ws.cell(row=row + 1, column=col, value=value)
    v.font = Font(size=16, bold=True, color=color or "FF0B0B0B")
    v.alignment = Alignment(horizontal="center", vertical="top")
    if fmt:
        v.number_format = fmt
    for r in (row, row + 1):
        for cc in range(col, col + span):
            ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=TILE_BG)


import re


def _month_from_name(stem):
    """MM/YYYY from an MM.YYYY / YYYY.MM pattern in the file name, else the stem."""
    m = re.search(r"(0?[1-9]|1[0-2])[._/](20\d{2})", stem) or \
        re.search(r"(20\d{2})[._/](0?[1-9]|1[0-2])", stem)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        mm, yy = (b, a) if a > 12 else (a, b)
        return f"{mm:02d}/{yy}"
    return stem.split("-", 1)[-1][:12] or stem[:12]


def pay_month_of(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    month = None
    for row in ws.iter_rows(min_row=1, max_row=12, values_only=True):
        for v in row:
            if isinstance(v, datetime):
                month = v
                break
        if month:
            break
    wb.close()
    return month


# Current-workbook rates (Progim 13.07.2026). When a slip legitimately matched
# an EARLIER official rate, that is a difference between the software and the
# workbook — the Progim only knows the current rate; the software carries the
# agreements' historical phase-ins.
CURRENT_RATES = {4544: 0.036, 4934: 0.05, 4994: 0.0725, 5401: 0.03875,
                 5216: 0.08, 741: 0.30, 5533: 0.05}


def _progim_delta(entry, raw_first, rules):
    """Why the software's verdict differs from a plain current-Progim run —
    one worker's notes, joined. Empty when the two agree exactly."""
    r = entry["result"]
    notes = []
    amt = defaultdict(float)
    for cp in r.components:
        amt[cp.code] += (cp.expected or 0.0)
    raw_vatek = float(raw_first[6] or 0)
    if abs((r.vatek_calculated or 0) - raw_vatek) > 1e-9:
        notes.append(f"ותק קטוע בקובץ: {raw_vatek} תוקן ל-{r.vatek_calculated} "
                     "(ה-Progim מניח ותק מדויק; הקובץ חותך לשתי ספרות)")
    raw_darga = str(raw_first[5] or "").strip()
    if raw_darga and r.darga_label and raw_darga != str(r.darga_label):
        notes.append(f"תווית דרגה נורמלה: '{raw_darga}' ← '{r.darga_label}' "
                     "(פלוס-לפני / קוד-דרגה; ה-Progim דורש תווית מדויקת)")
    if r.status == "valid":
        bs = sum(amt[c] for c in engine.BASE_CODES)
        bc = sum(cp.amount for cp in r.components if cp.calculated)
        if bc and abs(bc - bs) > 1.0:
            notes.append(f"הבסיס אושר דרך חלון עיגול הוותק (±0.125 שנה) — "
                         f"חישוב נקודתי כמו ב-Progim היה מראה פער {round(bc - bs, 2)} ₪")
    job = r.job_pct or 1.0
    for code, cur in CURRENT_RATES.items():
        rule = rules.get(code)
        if not rule or rule.get("type") != "percent":
            continue
        slip = sum(amt[c] for c in rule["codes"])
        if slip <= 0.01:
            continue
        base = sum(amt[c] for c in rule["base_codes"])
        base += rule.get("base_const", 0.0) * job
        if base <= 0:
            continue
        rate = slip / base
        matched = any(abs(base * rr - slip) <= 1.0 for rr in rule["rates"])
        if matched and abs(rate - cur) > 0.0015:
            notes.append(f"סמל {code}: שולם לפי אחוז תקופתי {rate * 100:.2f}% — "
                         f"בחוברת הנוכחית {cur * 100:.2f}% (ה-Progim מכיר רק את "
                         "האחוז הנוכחי; התוכנה מכירה את כל הפעימות)")
    return "; ".join(notes)


def _gap_reason(code, rules):
    """Short Hebrew 'why' for a gap on a pay code — for the per-code breakdown."""
    if code == "בסיס":
        return "שכר הבסיס אינו תואם — שינוי ותק/דרגה או חודש חלקי"
    if code in (667, 897):
        return "חריגה מהערך התקני לקבוצת הדרגה (חשד להפרשי רטרו)"
    t = (rules.get(code) or {}).get("type")
    if t == "shekel":
        return "סכום קבוע מחוץ לטבלת הערכים הרשמית"
    if t == "max22":
        return "חריגה מנוסחת 4550 (הגבוה מבין 22% למינימום)"
    if t == "percent":
        return "האחוז או בסיס-החישוב אינם תואמים את החוקה"
    if code in (5402, 5524):
        return "סכום שקלי החורג מנורמת קבוצת הדרגה בקובץ"
    return "סטייה מהערך התקני של הרכיב"


def collect(paths):
    """Run the engine per file; return (summary, per_emp, code_gaps)."""
    lookups = engine.get_lookups()
    rules = engine.get_rules()
    code_gaps = {}   # code -> {"name", "count", "sum"}
    files = sorted(paths, key=lambda p: (pay_month_of(p) or datetime(2099, 1, 1),
                                         Path(p).name))
    summary, per_emp = [], []
    for path in files:
        d = pay_month_of(path)
        month = d.strftime("%m/%Y") if d else _month_from_name(Path(path).stem)
        short = Path(path).stem.split("-", 1)[-1][:20] or Path(path).stem[:20]
        workers = engine.load_golmi(path)
        entries = engine.run_engine_full(workers, lookups)
        c = Counter(e["result"].status for e in entries)
        active = c["valid"] + c["invalid"]
        file_start = len(per_emp)   # breakdown is computed after the rows below
        summary.append({
            "month": month, "file": short, "workers": len(workers),
            "valid": c["valid"], "invalid": c["invalid"],
            "no_base": c["no_base"], "multi": c["multi_period"],
            "acc": round(c["valid"] / active * 100, 2) if active else 0.0,
        })
        print(f"  {month}  {len(workers):>7,} עובדים · שגויים {c['invalid']:>5,} "
              f"· {summary[-1]['acc']:.2f}%  ({short})")
        raw_firsts = [rows[0] for rows in workers.values()]
        for e, raw_first in zip(entries, raw_firsts):
            r, flags = e["result"], e["comp_flags"]
            bs = sum((cp.expected or 0.0) for cp in r.components
                     if cp.code in engine.BASE_CODES)
            bc = sum(cp.amount for cp in r.components if cp.calculated)
            # gap in גמולי השתלמות (slip − standard), split so A (667) and B
            # (897) sit next to the base gap; empty when the גמול matched.
            gmul_a = (round(flags[667]["slip"] - flags[667]["expected"], 2)
                      if 667 in flags else None)
            gmul_b = (round(flags[897]["slip"] - flags[897]["expected"], 2)
                      if 897 in flags else None)
            gap_798 = (round(flags[798]["slip"] - flags[798]["expected"], 2)
                       if 798 in flags else None)
            gap_4983 = (round(flags[4983]["slip"] - flags[4983]["expected"], 2)
                        if 4983 in flags else None)
            # result.total/total_diff already include every flagged component's
            # correction (folded in by the engine) — use them as-is.
            total_calc = round(r.total, 2)
            total_diff = round(r.total_diff or 0.0, 2)
            base_diff = round(bc - bs, 2) if bc else None
            # Error category, deduplicated by priority: a worker with several
            # gaps counts once, under the first bucket in this order — so
            # neutralizing each known cause leaves only the true unknowns.
            err_cat = None
            if r.status == "invalid":
                if base_diff is not None and abs(base_diff) > 1.0:
                    err_cat = "base"
                elif 667 in flags or 897 in flags:
                    err_cat = "gmul"
                elif 798 in flags:
                    err_cat = "brich"
                elif 4983 in flags:
                    err_cat = "mnhal"
                else:
                    err_cat = "real"
                # Per-code gap tally for the exec dashboard: base (when off) plus
                # every flagged component, each with its ₪ magnitude.
                def _bump(code, name, amt):
                    g = code_gaps.setdefault(code, {"name": name, "count": 0, "sum": 0.0})
                    g["count"] += 1
                    g["sum"] += abs(amt or 0.0)
                    if name:
                        g["name"] = name
                if base_diff is not None and abs(base_diff) > 1.0:
                    _bump("בסיס", "שכר בסיס", base_diff)
                for k, v in flags.items():
                    _bump(k, v["name"], v["slip"] - v["expected"])
            per_emp.append({
                "month": month, "file": short, "worker_id": r.worker_id,
                "ministry": r.ministry_name, "darga": r.darga_label,
                "vatek": r.vatek_calculated, "job_pct": r.job_pct,
                "full_time": (r.job_pct or 1.0) >= 0.999,
                "err_cat": err_cat,
                "base_slip": round(bs, 2),
                "base_calc": round(bc, 2) if bc else None,
                "base_diff": base_diff,
                "gmul_a": gmul_a, "gmul_b": gmul_b, "gap_798": gap_798,
                "gap_4983": gap_4983,
                "total_slip": r.expected_total, "total_calc": total_calc,
                "total_diff": total_diff, "status": r.status,
                "flags": "; ".join(
                    f"{k} ({v['name']}): {v['slip']} במקום {v['expected']}"
                    for k, v in sorted(flags.items())),
                "diag": "; ".join(r.errors),
                "progim_delta": _progim_delta(e, raw_first, rules),
            })
        # Full-time breakdown for the dashboard: part-timers are neutralized,
        # and each invalid full-timer lands in exactly ONE bucket (base >
        # gmul > דריכות > אמיתי) so the buckets never double-count.
        rows_f = per_emp[file_start:]
        ft = [x for x in rows_f if x["full_time"]]
        s = summary[-1]
        s["part_time"] = len(rows_f) - len(ft)
        s["ft"] = len(ft)
        s["ft_valid"] = sum(1 for x in ft if x["status"] == "valid")
        for cat, key in (("base", "inv_base"), ("gmul", "inv_gmul"),
                         ("brich", "inv_brich"), ("mnhal", "inv_mnhal"),
                         ("real", "inv_real")):
            s[key] = sum(1 for x in ft
                         if x["status"] == "invalid" and x["err_cat"] == cat)
        ft_active = s["ft_valid"] + sum(
            s[k] for k in ("inv_base", "inv_gmul", "inv_brich", "inv_mnhal", "inv_real"))
        s["real_pct"] = round(s["inv_real"] / ft_active * 100, 2) if ft_active else 0.0
    code_gap_list = sorted(
        ({"code": code, "name": g["name"], "count": g["count"],
          "sum": round(g["sum"], 2), "reason": _gap_reason(code, rules)}
         for code, g in code_gaps.items()),
        key=lambda x: -x["count"])
    return summary, per_emp, code_gap_list


# (key, header, width, number-format, is-gap-cell). Gap cells are red-tinted
# when they carry a real value so the eye lands on them.
EMP_COLS = [
    ("month", "חודש שכר", 11, None, False), ("file", "קובץ", 18, None, False),
    ("worker_id", "מסד עובד", 12, INT, False), ("ministry", "משרד", 22, None, False),
    ("darga", "דרגה", 8, None, False), ("vatek", "ותק", 8, None, False),
    ("job_pct", "חלקיות", 8, None, False), ("base_slip", "בסיס בתלוש", 13, MONEY, False),
    ("base_calc", "בסיס מחושב", 13, MONEY, False), ("base_diff", "הפרש בסיס", 12, MONEY, True),
    ("gmul_a", "פער גמול א'", 12, MONEY, True), ("gmul_b", "פער גמול ב'", 12, MONEY, True),
    ("gap_798", "פער דריכות בי\"ח", 13, MONEY, True),
    ("gap_4983", "פער גמול מנהל", 12, MONEY, True),
    ("total_slip", "סכום בתלוש", 13, MONEY, False), ("total_calc", "סכום מחושב", 13, MONEY, False),
    ("total_diff", "הפרש כולל", 12, MONEY, True), ("status_he", "סטטוס", 16, None, False),
    ("flags", "רכיבים חריגים", 30, None, False), ("diag", "אבחון", 30, None, False),
    ("progim_delta", "שוני מול Progim — והסבר", 42, None, False),
]


def _emp_sheet(wb, title, rows, table_name, highlight_invalid):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = "D2"            # keep month/file/מסד visible when scrolling ₪ cols
    _header_row(ws, 1, [he for _, he, _, _, _ in EMP_COLS],
                [w for _, _, w, _, _ in EMP_COLS])
    inv_font = Font(color=BAD_TXT, bold=True)
    inv_fill = PatternFill("solid", fgColor=BAD_BG)
    ok_font = Font(color=GOOD_TXT)
    warn_font = Font(color=WARN_TXT)
    for r_i, row in enumerate(rows, start=2):
        vals = [row.get(k) if k != "status_he" else STATUS_HE[row["status"]]
                for k, _, _, _, _ in EMP_COLS]
        for c_i, ((key, _, _, fmt, is_gap), v) in enumerate(zip(EMP_COLS, vals), start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            if fmt:
                cell.number_format = fmt
            if key == "status_he":
                st = row["status"]
                cell.font = (inv_font if st == "invalid"
                             else ok_font if st == "valid" else warn_font)
                if st == "invalid":
                    cell.fill = inv_fill
            elif is_gap and isinstance(v, (int, float)) and abs(v) > 1:
                # any real gap gets a red tint — the eye lands on it directly
                cell.font = inv_font
                cell.fill = inv_fill
        if highlight_invalid and row["status"] == "invalid":
            ws.cell(row=r_i, column=1).fill = inv_fill
    if rows:
        ref = f"A1:{get_column_letter(len(EMP_COLS))}{len(rows) + 1}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight15",
                                              showRowStripes=True)
        ws.add_table(table)
    return ws


def _anomaly(row):
    """₪ magnitude of a row's gap: larger of the base gap and the corrected-total
    gap (the latter folds in component corrections incl. גמול), so a component-
    only mismatch still ranks."""
    return max(abs(row["base_diff"] or 0.0), abs(row["total_diff"] or 0.0))


def _month_key(m):
    try:
        mm, yy = str(m).split("/")
        return int(yy) * 12 + int(mm)
    except ValueError:
        return 0


def compute_flips(per_emp):
    """Workers whose slip flipped valid→invalid between consecutive months."""
    hist = defaultdict(list)
    for r in per_emp:
        hist[r["worker_id"]].append(r)
    flips = []
    for rows in hist.values():
        if len(rows) < 2:
            continue
        rows.sort(key=lambda r: _month_key(r["month"]))
        for prev, cur in zip(rows, rows[1:]):
            if (prev["status"] == "valid" and cur["status"] == "invalid"
                    and _month_key(prev["month"]) != _month_key(cur["month"])):
                flips.append({"id": cur["worker_id"], "ministry": cur["ministry"],
                              "darga": cur["darga"], "from": prev["month"],
                              "to": cur["month"], "diff": _anomaly(cur)})
    return sorted(flips, key=lambda f: -f["diff"])


def write_workbook(summary, per_emp, out_path, code_gaps=None):
    wb = openpyxl.Workbook()
    anom_by_file = defaultdict(float)
    anom_total = 0.0
    for r in per_emp:
        if r["status"] == "invalid":
            v = _anomaly(r)
            anom_by_file[(r["month"], r["file"])] += v
            anom_total += v

    # ---- לוח בקרה ------------------------------------------------------------
    ws = wb.active
    ws.title = "לוח בקרה"
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:N1")
    t = ws["A1"]
    t.value = "בדיקת התאמת תלושים — דוח מאוחד"
    t.font = Font(size=16, bold=True, color=NAVY)
    ws.merge_cells("A2:N2")
    s = ws["A2"]
    s.value = (f"{len(summary)} קבצים · הופק "
               f"{datetime.now().strftime('%d/%m/%Y %H:%M')} · "
               "מנטרל: משרה חלקית ← בסיס ← גמול השתלמות ← דריכות בי\"ח ← "
               "גמול מנהל. הנותר = שגיאה אמיתית לתיקון. (רטרו רב-תקופתי = "
               "מחוץ לשרשרת)")
    s.font = Font(size=10, color=MUTED)

    tot = Counter()
    for r in summary:
        for k in ("workers", "valid", "invalid", "no_base", "multi"):
            tot[k] += r[k]
    active = tot["valid"] + tot["invalid"]
    acc = round(tot["valid"] / active * 100, 2) if active else 0.0
    _kpi(ws, 4, 1, 1, "סה\"כ עובדים", tot["workers"], fmt=INT)
    _kpi(ws, 4, 2, 1, "תקינים", tot["valid"], GOOD_TXT, INT)
    for k in ("part_time", "ft", "ft_valid", "inv_base", "inv_gmul",
              "inv_brich", "inv_mnhal", "inv_real"):
        tot[k] = sum(r.get(k, 0) for r in summary)
    _ft_active = (tot["ft_valid"] + tot["inv_base"] + tot["inv_gmul"]
                  + tot["inv_brich"] + tot["inv_mnhal"] + tot["inv_real"])
    _real_pct = (tot["inv_real"] / _ft_active) if _ft_active else 0.0
    _kpi(ws, 4, 3, 1, "שגויים אמיתיים (מלאה)", tot["inv_real"], BAD_TXT, INT)
    _kpi(ws, 4, 4, 1, "% שגויים אמיתיים", _real_pct,
         GOOD_TXT if _real_pct <= 0.01 else BAD_TXT, "0.00%")
    _kpi(ws, 4, 5, 1, "שווי חריגות ₪", round(anom_total), BAD_TXT, INT)
    _kpi(ws, 4, 6, 1, "ללא בסיס פעיל", tot["no_base"], WARN_TXT, INT)
    _kpi(ws, 4, 7, 1, "רטרו / רב-תקופתי", tot["multi"], WARN_TXT, INT)

    head_r = 7
    # Neutralization chain (each invalid full-timer counts once, in this order):
    # base → gmul → דריכות → גמול מנהל. What remains is the REAL error count.
    # רטרו/רב-תקופתי is a SEPARATE class (base code appears twice — the slip
    # merges two pay periods), not part of the chain, shown for reference only.
    labels = ["חודש שכר", "קובץ", "עובדים", "משרה חלקית", "משרה מלאה", "תקין (מלאה)",
              "שגויי בסיס", "שגויי גמול", "שגויי דריכות", "שגויי גמול מנהל",
              "שגויים אמיתיים", "ללא בסיס", "רטרו רב-תקופתי", "% שגויים אמיתיים"]
    _header_row(ws, head_r, labels,
                [11, 18, 10, 10, 10, 11, 10, 10, 10, 11, 12, 9, 11, 13])
    for i, r in enumerate(summary, start=head_r + 1):
        vals = [r["month"], r["file"], r["workers"], r.get("part_time", 0),
                r.get("ft", 0), r.get("ft_valid", 0), r.get("inv_base", 0),
                r.get("inv_gmul", 0), r.get("inv_brich", 0), r.get("inv_mnhal", 0),
                r.get("inv_real", 0), r["no_base"], r["multi"],
                r.get("real_pct", 0.0) / 100]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=i, column=c_i, value=v)
            cell.border = THIN_BOX
            if 3 <= c_i <= 13:
                cell.number_format = INT
            if c_i == 6:
                cell.font = Font(color=GOOD_TXT)
            if c_i in (7, 8, 9, 10) and v:
                cell.font = Font(color=WARN_TXT)
            if c_i == 11 and v:
                cell.font = Font(color=BAD_TXT, bold=True)
            if c_i == 14:
                cell.number_format = "0.00%"
    last = head_r + len(summary)
    trow = last + 1
    tot_active = _ft_active
    real_pct_tot = _real_pct
    tvals = ["סה\"כ", "", tot["workers"], tot["part_time"], tot["ft"],
             tot["ft_valid"], tot["inv_base"], tot["inv_gmul"], tot["inv_brich"],
             tot["inv_mnhal"], tot["inv_real"], tot["no_base"], tot["multi"],
             real_pct_tot]
    for c_i, v in enumerate(tvals, start=1):
        cell = ws.cell(row=trow, column=c_i, value=v)
        cell.font = Font(bold=True)
        cell.border = Border(top=Side(style="double", color=NAVY))
        if 3 <= c_i <= 13:
            cell.number_format = INT
        if c_i == 14:
            cell.number_format = "0.00%"
        if c_i == 11:
            cell.font = Font(bold=True, color=BAD_TXT)
    rng = f"N{head_r + 1}:N{last}"    # % שגויים אמיתיים
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="lessThanOrEqual", formula=["0.01"],
        font=Font(color=GOOD_TXT), fill=PatternFill("solid", fgColor=GOOD_BG)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="between", formula=["0.0101", "0.05"],
        font=Font(color=WARN_TXT), fill=PatternFill("solid", fgColor=WARN_BG)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="greaterThan", formula=["0.05"],
        font=Font(color=BAD_TXT), fill=PatternFill("solid", fgColor=BAD_BG)))
    ws.conditional_formatting.add(
        f"C{head_r + 1}:C{last}",
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color=BAR_BLUE, showValue=True))
    ws.conditional_formatting.add(
        f"K{head_r + 1}:K{last}",          # שגויים אמיתיים
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color="FFD03B3B", showValue=True))

    # ---- פערים לפי סמל שכר (a second table, to the right of the per-file one) ---
    if code_gaps:
        cbase = 16   # column P (leaves a gap after the 14-column per-file table)
        heads = ["סמל", "שם רכיב", "כמות פערים", "שווי ₪", "הסיבה לפער"]
        widths = [9, 22, 12, 13, 40]
        for j, (h, w) in enumerate(zip(heads, widths)):
            c = ws.cell(row=head_r, column=cbase + j, value=h)
            c.font = Font(bold=True, color=NAVY_TEXT, size=11)
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = THIN_BOX
            ws.column_dimensions[get_column_letter(cbase + j)].width = w
        cl = ws.cell(row=head_r - 1, column=cbase, value="פערים לפי סמל שכר — כמה וסיבה")
        cl.font = Font(bold=True, size=12, color=NAVY)
        for gi, g in enumerate(code_gaps, start=head_r + 1):
            vals = [g["code"], g["name"], g["count"], round(g["sum"]), g["reason"]]
            for j, v in enumerate(vals):
                cell = ws.cell(row=gi, column=cbase + j, value=v)
                cell.border = THIN_BOX
                if j in (2, 3):
                    cell.number_format = INT
                if j == 2 and g["count"]:
                    cell.font = Font(bold=True, color=BAD_TXT)
        gl = head_r + len(code_gaps)
        ws.conditional_formatting.add(
            f"{get_column_letter(cbase + 2)}{head_r + 1}:{get_column_letter(cbase + 2)}{gl}",
            DataBarRule(start_type="num", start_value=0, end_type="max",
                        color="FFD03B3B", showValue=True))

    # ---- שגויים לבדיקה (ממוין לפי ₪) --------------------------------------------
    inv = [r for r in per_emp if r["status"] == "invalid"
           and r["full_time"] and r["err_cat"] == "real"]
    inv.sort(key=_anomaly, reverse=True)
    _emp_sheet(wb, "שגויים לבדיקה", inv, "Invalids", highlight_invalid=False)

    # ---- שינויי סטטוס בין חודשים -------------------------------------------------
    flips = compute_flips(per_emp)
    if flips:
        wsf = wb.create_sheet("שינויי סטטוס")
        wsf.sheet_view.rightToLeft = True
        wsf.freeze_panes = "A2"
        _header_row(wsf, 1, ["מסד עובד", "משרד", "דרגה", "תקין בחודש",
                             "שגוי בחודש", "פער ₪"], [12, 22, 8, 12, 12, 12])
        bad = Font(color=BAD_TXT, bold=True)
        for i, f in enumerate(flips, start=2):
            vals = [f["id"], f["ministry"], f["darga"], f["from"], f["to"], f["diff"]]
            for c_i, v in enumerate(vals, start=1):
                cell = wsf.cell(row=i, column=c_i, value=v)
                cell.border = THIN_BOX
            wsf.cell(row=i, column=1).number_format = INT
            wsf.cell(row=i, column=6).number_format = MONEY
            wsf.cell(row=i, column=5).font = bad
            wsf.cell(row=i, column=6).font = bad
        wsf.auto_filter.ref = f"A1:F{len(flips) + 1}"

    # ---- פר עובד ---------------------------------------------------------------
    _emp_sheet(wb, "פר עובד", per_emp, "PerEmployee", highlight_invalid=True)

    # ---- פילוח משרדים ----------------------------------------------------------
    agg = defaultdict(lambda: Counter())
    for r in per_emp:
        agg[r["ministry"] or "—"][r["status"]] += 1
    ws4 = wb.create_sheet("פילוח משרדים")
    ws4.sheet_view.rightToLeft = True
    ws4.freeze_panes = "A2"
    _header_row(ws4, 1, ["משרד / גוף", "עובדים", "תקין", "שגוי", "% תקינות (פעילים)"],
                [26, 11, 11, 9, 16])
    rows = sorted(agg.items(), key=lambda kv: -sum(kv[1].values()))
    for i, (name, c) in enumerate(rows, start=2):
        act = c["valid"] + c["invalid"]
        vals = [name, sum(c.values()), c["valid"], c["invalid"],
                (c["valid"] / act) if act else None]
        for c_i, v in enumerate(vals, start=1):
            cell = ws4.cell(row=i, column=c_i, value=v)
            cell.border = THIN_BOX
            if c_i in (2, 3, 4):
                cell.number_format = INT
            if c_i == 4 and c["invalid"]:
                cell.font = Font(color=BAD_TXT, bold=True)
            if c_i == 5:
                cell.number_format = "0.0%"
    last4 = len(rows) + 1
    ws4.conditional_formatting.add(
        f"B2:B{last4}", DataBarRule(start_type="num", start_value=0,
                                    end_type="max", color=BAR_BLUE, showValue=True))
    rng4 = f"E2:E{last4}"
    ws4.conditional_formatting.add(rng4, CellIsRule(
        operator="greaterThanOrEqual", formula=["0.99"],
        fill=PatternFill("solid", fgColor=GOOD_BG)))
    ws4.conditional_formatting.add(rng4, CellIsRule(
        operator="lessThan", formula=["0.95"],
        fill=PatternFill("solid", fgColor=BAD_BG)))
    tbl4 = Table(displayName="Ministries", ref=f"A1:E{last4}")
    tbl4.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=True)
    ws4.add_table(tbl4)

    # ---- מגמה בין-חודשית: native Excel charts on the dashboard -----------------
    months = sorted({r["month"] for r in summary},
                    key=lambda m: (m.split("/")[-1], m.split("/")[0]))
    if len(months) > 1:
        per_month = {m: Counter() for m in months}
        for r in summary:
            for k in ("workers", "valid", "invalid"):
                per_month[r["month"]][k] += r[k]
        anchor = trow + 3
        ws.cell(row=anchor - 1, column=1, value="מגמה בין-חודשית").font = \
            Font(bold=True, color=NAVY)
        _header_row(ws, anchor, ["חודש", "% תקינות", "שגויים"], [11, 11, 9])
        for i, m in enumerate(months, start=anchor + 1):
            c = per_month[m]
            act = c["valid"] + c["invalid"]
            ws.cell(row=i, column=1, value=m).border = THIN_BOX
            pc = ws.cell(row=i, column=2, value=(c["valid"] / act) if act else None)
            pc.number_format = "0.00%"
            pc.border = THIN_BOX
            ic = ws.cell(row=i, column=3, value=c["invalid"])
            ic.number_format = INT
            ic.border = THIN_BOX
        from openpyxl.chart import BarChart, LineChart, Reference
        lastm = anchor + len(months)
        line = LineChart()
        line.title = "% תקינות לפי חודש"
        line.height, line.width = 7, 13
        line.y_axis.numFmt = "0%"
        line.add_data(Reference(ws, min_col=2, min_row=anchor, max_row=lastm),
                      titles_from_data=True)
        line.set_categories(Reference(ws, min_col=1, min_row=anchor + 1, max_row=lastm))
        ws.add_chart(line, f"E{anchor - 1}")
        bar = BarChart()
        bar.title = "שגויים לפי חודש"
        bar.height, bar.width = 7, 13
        bar.add_data(Reference(ws, min_col=3, min_row=anchor, max_row=lastm),
                     titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=1, min_row=anchor + 1, max_row=lastm))
        ws.add_chart(bar, f"E{anchor + 15}")

        # ---- שינויים בין חודשים: עובד תקין שהפך שגוי ----------------------------
        month_idx = {m: i for i, m in enumerate(months)}
        by_worker = defaultdict(dict)
        for r in per_emp:
            if r["status"] in ("valid", "invalid"):
                by_worker[r["worker_id"]][r["month"]] = r
        regressions = []
        for wid, per_m in by_worker.items():
            ms = sorted(per_m, key=lambda m: month_idx.get(m, 99))
            for a, b in zip(ms, ms[1:]):
                if per_m[a]["status"] == "valid" and per_m[b]["status"] == "invalid":
                    cur = per_m[b]
                    regressions.append({
                        "worker_id": wid, "ministry": cur["ministry"],
                        "darga": cur["darga"], "from_m": a, "to_m": b,
                        "total_diff": cur["total_diff"], "flags": cur["flags"],
                        "diag": cur["diag"],
                    })
        ws5 = wb.create_sheet("שינויים בין חודשים")
        ws5.sheet_view.rightToLeft = True
        ws5.freeze_panes = "A3"
        ws5.merge_cells("A1:H1")
        h = ws5["A1"]
        h.value = ("עובדים שהיו תקינים בחודש מוקדם והפכו שגויים בחודש מאוחר — "
                   f"{len(regressions)} מקרים · ממוין לפי גודל הפער")
        h.font = Font(bold=True, color=NAVY)
        _header_row(ws5, 2, ["מסד עובד", "משרד", "דרגה", "תקין בחודש",
                             "שגוי בחודש", "הפרש כולל", "רכיבים חריגים", "אבחון"],
                    [12, 22, 8, 12, 12, 12, 34, 34])
        regressions.sort(key=lambda r: abs(r["total_diff"] or 0), reverse=True)
        for i, r in enumerate(regressions, start=3):
            vals = [r["worker_id"], r["ministry"], r["darga"], r["from_m"],
                    r["to_m"], r["total_diff"], r["flags"], r["diag"]]
            for c_i, v in enumerate(vals, start=1):
                cell = ws5.cell(row=i, column=c_i, value=v)
                cell.border = THIN_BOX
                if c_i == 1:
                    cell.number_format = INT
                if c_i == 6:
                    cell.number_format = MONEY
                    cell.font = Font(color=BAD_TXT, bold=True)

    wb.save(out_path)


def main_cli():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("files", nargs="+", help="גולמי .xlsx files")
    ap.add_argument("--out", default="unified.xlsx")
    args = ap.parse_args()
    t0 = time.time()
    summary, per_emp, code_gaps = collect(args.files)
    print(f"עיבוד: {time.time() - t0:.0f}ש · כותב workbook ({len(per_emp):,} שורות)...")
    write_workbook(summary, per_emp, args.out, code_gaps)
    inv = sum(1 for r in per_emp if r["status"] == "invalid")
    print(f"נכתב: {args.out} · {len(per_emp):,} רשומות · {inv:,} שגויים")


if __name__ == "__main__":
    main_cli()
