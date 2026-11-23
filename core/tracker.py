"""The cohort tracker: one sheet per cohort, refreshed in place, plus a change log.
New people are appended, changes are updated and logged, people who leave the report
are kept and flagged, and your own columns are never overwritten."""
import os
import shutil
import sys
from collections import Counter

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .rules import key, now, phone_fmt, sort_key, txt, us_date, zip5, is_ship_ready, is_do_not_ship

SYS = ["Full Name", "Status", "Street Address Line 1", "Street Address Line 2",
       "City", "State", "Zip Code", "Phone Number", "Recruiter Notes"]
META = ["Row State", "First Seen", "Last Updated"]
MANUAL = ["Shipped?", "Ship Date", "Tracking Number", "Internal Notes"]
HEADERS = SYS + META + MANUAL
WATCH = ["Status", "Street Address Line 1", "Street Address Line 2", "City",
         "State", "Zip Code", "Phone Number", "Recruiter Notes"]
LOG_SHEET = "Change Log"
LOG_HEADERS = ["Timestamp", "Cohort", "Full Name", "Change", "Field",
               "Previous Value", "New Value"]
SRC_COHORT = "Cohort Date"

WIDTHS = [26, 24, 30, 18, 20, 8, 10, 17, 52,   # report columns (full-roster layout)
          20, 20, 20,                           # Row State, First Seen, Last Updated
          10, 12, 18, 30]                       # your columns
CENTERED = {"State", "Zip Code", "Row State"}
WRAPPED = {"Recruiter Notes", "Internal Notes"}
THIN = Side(style="thin", color="BFBFBF")
HDR_FILL = PatternFill("solid", fgColor="1F4E79")


# ---------------------------------------------------------------- read
def cohort_records(df, cohort):
    import pandas as pd
    for c in SYS + [SRC_COHORT]:
        if c not in df.columns:
            sys.exit(f"Report is missing column: {c}")
    sub = df[df[SRC_COHORT] == pd.Timestamp(cohort)]
    if sub.empty:
        avail = ", ".join(sorted({str(d)[:10] for d in df[SRC_COHORT].dropna()}))
        sys.exit(f"No rows for cohort {cohort}.\nAvailable cohorts: {avail}")
    out = {}
    for _, r in sub.iterrows():
        rec = {c: txt(r[c]) for c in SYS}
        rec["Zip Code"] = zip5(r["Zip Code"])
        rec["Phone Number"] = phone_fmt(r["Phone Number"])
        k = key(rec["Full Name"])
        if k in out:
            print(f"  warning: {rec['Full Name']} appears twice in the report; keeping the first")
            continue
        out[k] = rec
    return out


def header_row(ws):
    for r in range(1, min(ws.max_row, 12) + 1):
        if txt(ws.cell(row=r, column=1).value) == "Full Name":
            return r
    return None


def read_sheet(wb, sheet):
    if wb is None or sheet not in wb.sheetnames:
        return {}
    ws = wb[sheet]
    hr = header_row(ws)
    if hr is None:
        return {}
    cols = {txt(ws.cell(row=hr, column=c).value): c
            for c in range(1, ws.max_column + 1) if txt(ws.cell(row=hr, column=c).value)}
    rows = {}
    for r in range(hr + 1, ws.max_row + 1):
        name = txt(ws.cell(row=r, column=cols["Full Name"]).value)
        if name:
            rows[key(name)] = {h: txt(ws.cell(row=r, column=c).value) for h, c in cols.items()}
    return rows


def cohort_sheets(wb):
    import re
    return sorted(s for s in wb.sheetnames if re.fullmatch(r"\d{4}-\d\d-\d\d", s)) if wb else []


# ---------------------------------------------------------------- merge
def merge(existing, report, cohort):
    stamp = now()
    merged, log = {}, []
    for k, rec in report.items():
        prev = existing.get(k)
        row = dict(rec)
        if prev is None:
            row["Row State"] = "New"
            row["First Seen"] = stamp
            row["Last Updated"] = stamp
            for m in MANUAL:
                row[m] = ""
            log.append([stamp, cohort, rec["Full Name"], "Added", "", "", rec["Status"]])
        else:
            diffs = [(f, prev.get(f, ""), rec[f]) for f in WATCH if prev.get(f, "") != rec[f]]
            for m in MANUAL:
                row[m] = prev.get(m, "")
            row["First Seen"] = prev.get("First Seen") or stamp
            if diffs:
                row["Row State"] = "Updated"
                row["Last Updated"] = stamp
                for f, old, new in diffs:
                    log.append([stamp, cohort, rec["Full Name"], "Changed", f, old, new])
            else:
                row["Row State"] = "Unchanged"
                row["Last Updated"] = prev.get("Last Updated") or stamp
        merged[k] = row
    for k, prev in existing.items():
        if k in report:
            continue
        row = dict(prev)
        if prev.get("Row State") != "Removed from report":
            row["Row State"] = "Removed from report"
            row["Last Updated"] = stamp
            log.append([stamp, cohort, prev.get("Full Name", ""), "Removed from report",
                        "Status", prev.get("Status", ""), ""])
        merged[k] = row
    return merged, log


# ---------------------------------------------------------------- write
def fill_for(cfg, row):
    if row.get("Row State") == "Removed from report":
        return "F8CBAD"
    if is_ship_ready(cfg, row.get("Status")):
        return "E2EFDA"
    if is_do_not_ship(cfg, row.get("Status")):
        return "D9D9D9"
    return "FFF2CC"


def write_sheet(cfg, wb, cohort, merged):
    sheet = cohort
    if sheet in wb.sheetnames:
        del wb[sheet]
    ws = wb.create_sheet(sheet)
    try:                                   # keep cohorts in date order
        wb.move_sheet(sheet, offset=-(len(wb.sheetnames) - 1 -
                      sorted(s for s in wb.sheetnames if s != LOG_SHEET).index(sheet)))
    except Exception:
        pass

    sk = lambda r: sort_key(cfg, r)
    rows = sorted(merged.values(), key=sk)
    live = [r for r in rows if r.get("Row State") != "Removed from report"]
    gone = len(rows) - len(live)
    ship = sum(1 for r in live if is_ship_ready(cfg, r.get("Status")))
    new = sum(1 for r in rows if r.get("Row State") == "New")
    upd = sum(1 for r in rows if r.get("Row State") == "Updated")
    counts = Counter(txt(r.get("Status")) for r in live)
    ordered = sorted((k for k in counts if k), key=lambda k: sk({"Status": k, "Full Name": ""}))
    activity = f"{new} added, {upd} updated" if (new or upd) else "no changes"
    notes = [
        (f"{us_date(cohort)} Cohort — Full Roster", "title"),
        (f"All {len(live)} names in the cohort, no status filter."
         + (f"  {gone} no longer in the report." if gone else "")
         + f"  Last run {now()}: {activity}.", "sub"),
        ("Sorted by status (Ready To Ship first), then name. Zips padded to 5 digits, "
         "phones reformatted. Addresses as-entered. Columns from Shipped? onward are "
         "yours and are never overwritten.", "sub"),
        ("GREEN = Ready To Ship.  YELLOW = pending/on hold.  GREY = do not ship "
         "(already shipped, started, or dropped).  ORANGE = no longer in the report.", "legend"),
        ("  |  ".join(f"{k}: {counts[k]}" for k in ordered), "counts"),
    ]
    style = {"title": dict(size=13, bold=True, italic=False, color=None),
             "sub": dict(size=9, bold=False, italic=True, color="555555"),
             "legend": dict(size=9, bold=False, italic=True, color="C00000"),
             "counts": dict(size=9, bold=True, italic=False, color="333333")}
    for i, (line, kind) in enumerate(notes, start=1):
        s = style[kind]
        ws.cell(row=i, column=1, value=line).font = Font(
            name="Arial", size=s["size"], bold=s["bold"], italic=s["italic"], color=s["color"])

    hr = len(notes) + 2
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(row=hr, column=i, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(THIN, THIN, THIN, THIN)
    for ri, row in enumerate(rows):
        bg = fill_for(cfg, row)
        flagged = row.get("Row State") in ("New", "Updated", "Removed from report")
        for ci, h in enumerate(HEADERS, 1):
            c = ws.cell(row=hr + 1 + ri, column=ci, value=row.get(h, ""))
            c.font = Font(name="Arial", size=10, bold=(h == "Row State" and flagged))
            c.border = Border(THIN, THIN, THIN, THIN)
            c.alignment = Alignment(horizontal="center" if h in CENTERED else "left",
                                    vertical="top" if h in WRAPPED else "center",
                                    wrap_text=(h in WRAPPED))
            if h == "Zip Code":
                c.number_format = "@"
            c.fill = PatternFill("solid", fgColor=bg)
    for ci, w in enumerate(WIDTHS[:len(HEADERS)], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[hr].height = 28
    ws.freeze_panes = f"A{hr + 1}"
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(HEADERS))}{hr + len(rows)}"
    return len(rows), ship, new, upd


def append_change_log(wb, entries):
    if LOG_SHEET in wb.sheetnames:
        ws = wb[LOG_SHEET]
    else:
        ws = wb.create_sheet(LOG_SHEET)
        for i, h in enumerate(LOG_HEADERS, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = HDR_FILL
            c.alignment = Alignment(horizontal="center", vertical="center")
        for ci, w in enumerate([17, 12, 26, 20, 24, 30, 30], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "A2"
    start = ws.max_row + 1
    for i, e in enumerate(entries):
        for ci, v in enumerate(e, 1):
            c = ws.cell(row=start + i, column=ci, value=v)
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(vertical="center", wrap_text=(ci >= 6))
    if entries:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(LOG_HEADERS))}{ws.max_row}"


def keep_change_log_last(wb):
    if LOG_SHEET in wb.sheetnames:
        wb.move_sheet(LOG_SHEET, offset=len(wb.sheetnames) - 1 - wb.sheetnames.index(LOG_SHEET))


# ---------------------------------------------------------------- workbook I/O
def open_tracker(cfg):
    return openpyxl.load_workbook(cfg.tracker) if os.path.exists(cfg.tracker) else None


def assert_writable(path):
    """Excel locks open workbooks. Find out before anything is written."""
    if not os.path.exists(path):
        return
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        sys.exit(f"{os.path.basename(path)} is open in Excel (or read-only). "
                 f"Close it and run again.\nNothing was written.")


def save_tracker(cfg, wb, backup=True):
    if backup and os.path.exists(cfg.tracker):
        shutil.copy2(cfg.tracker, cfg.tracker + ".bak")
    keep_change_log_last(wb)
    wb.save(cfg.tracker)


# ---------------------------------------------------------------- the operation
def refresh(cfg, df, cohort, backup=True):
    """Bring one cohort's sheet up to date from the report DataFrame."""
    report = cohort_records(df, cohort)
    assert_writable(cfg.tracker)
    wb = open_tracker(cfg)
    existing = read_sheet(wb, cohort)
    merged, run_log = merge(existing, report, cohort)
    if wb is None:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
    total, ship, new, upd = write_sheet(cfg, wb, cohort, merged)
    append_change_log(wb, run_log)
    save_tracker(cfg, wb, backup=backup)

    name = os.path.basename(cfg.tracker)
    print(f"{'Created' if not existing else 'Updated'} {name} [{cohort}]")
    print(f"  tracked: {total}   ready to ship: {ship}")
    if run_log:
        print(f"  this run: {new} added, {upd} updated, "
              f"{sum(1 for e in run_log if e[3] == 'Removed from report')} removed")
        for e in run_log[:25]:
            detail = (f"{e[4]}: {e[5] or '(blank)'} -> {e[6] or '(blank)'}"
                      if e[3] == "Changed" else e[3])
            print(f"    {e[2]}: {detail}")
        if len(run_log) > 25:
            print(f"    ... and {len(run_log) - 25} more (see {LOG_SHEET})")
    else:
        print("  this run: no changes")
    return run_log
