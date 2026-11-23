"""FedEx label batches. Each batch is a fresh copy of the template, and everyone in it is
recorded in the Label Log tab of the tracker, so nobody is ever labeled twice."""
import os
import re
import sys
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import tracker as T
from .rules import (ADDRESS_FIELDS, address_problems, is_ship_ready, key, now, phone_fmt,
                    same_address, txt, us_date, zip5)
from .ui import ask, confirm

LOG_SHEET = "Label Log"
LOG_HEADERS = ["Created", "Cohort", "Batch", "Full Name", "Status at Label",
               "Street Address Line 1", "Street Address Line 2", "City", "State",
               "Zip Code", "Phone Number", "Batch File"]
RECIPIENT = {                       # FedEx template header -> tracker column
    "recipientContactName": "Full Name",
    "recipientContactNumber": "Phone Number",
    "recipientLine1": "Street Address Line 1",
    "recipientLine2": "Street Address Line 2",
    "recipientPostcode": "Zip Code",
    "recipientCity": "City",
    "recipientState": "State",
}
BLANK_PER_RECIPIENT = {"recipientCompany"}


# ---------------------------------------------------------------- tracker access
def load_cohort(cfg, cohort):
    if not os.path.exists(cfg.tracker):
        sys.exit(f"{os.path.basename(cfg.tracker)} not found. Refresh from Smartsheet first.")
    wb = openpyxl.load_workbook(cfg.tracker)
    if cohort not in wb.sheetnames:
        have = ", ".join(T.cohort_sheets(wb)) or "none"
        sys.exit(f"No {cohort} sheet in the tracker (has: {have}). Refresh that cohort first.")
    ws = wb[cohort]
    hr = T.header_row(ws)
    if hr is None:
        sys.exit(f"The {cohort} sheet has no header row. Refresh it again.")
    cols = {txt(ws.cell(row=hr, column=c).value): c
            for c in range(1, ws.max_column + 1) if ws.cell(row=hr, column=c).value}
    rows = [{h: txt(ws.cell(row=r, column=c).value) for h, c in cols.items()}
            for r in range(hr + 1, ws.max_row + 1)
            if txt(ws.cell(row=r, column=cols["Full Name"]).value)]
    m = re.search(r"Last run (\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)", txt(ws["A2"].value))
    last_run = datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S") if m else None
    return wb, rows, last_run


def read_log(wb):
    if LOG_SHEET not in wb.sheetnames:
        return []
    ws = wb[LOG_SHEET]
    hdr = [txt(ws.cell(row=1, column=c).value) for c in range(1, ws.max_column + 1)]
    out = []
    for r in range(2, ws.max_row + 1):
        rec = {h: txt(ws.cell(row=r, column=c).value) for c, h in enumerate(hdr, 1) if h}
        if rec.get("Full Name"):
            rec["_row"] = r
            out.append(rec)
    return out


def _log_sheet(wb):
    if LOG_SHEET in wb.sheetnames:
        return wb[LOG_SHEET]
    ws = wb.create_sheet(LOG_SHEET)
    for i, h in enumerate(LOG_HEADERS, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, w in enumerate([19, 12, 7, 26, 16, 30, 18, 18, 7, 10, 16, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    return ws


def _append_log(wb, cohort, batch, people, batch_file):
    ws = _log_sheet(wb)
    stamp = now()
    for p in people:
        vals = [stamp, cohort, batch, p["Full Name"], p.get("Status", ""),
                *[p.get(f, "") for f in ADDRESS_FIELDS], batch_file]
        r = ws.max_row + 1
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="Arial", size=10)
            if LOG_HEADERS[c - 1] == "Zip Code":
                cell.number_format = "@"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(LOG_HEADERS))}{ws.max_row}"


def _batch_numbers(log, cohort):
    return sorted({int(e["Batch"]) for e in log
                   if e.get("Cohort") == cohort and e.get("Batch", "").isdigit()})


# ---------------------------------------------------------------- template
def _template_columns(ws):
    hdr = {txt(ws.cell(row=1, column=c).value): c
           for c in range(1, ws.max_column + 1) if ws.cell(row=1, column=c).value}
    missing = [h for h in RECIPIENT if h not in hdr]
    if missing:
        sys.exit(f"The template is missing these columns: {', '.join(missing)}")
    return hdr


def _template(path):
    if not os.path.exists(path):
        sys.exit(f"Template not found: {path}")
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    hdr = _template_columns(ws)
    sender = hdr.get("senderContactName")
    model = next((r for r in range(2, ws.max_row + 1)
                  if sender and txt(ws.cell(row=r, column=sender).value)), None)
    if model is None:
        sys.exit("The template has no filled-in row to copy sender and package settings "
                 "from. Keep at least one complete row in it.")
    return wb, ws, hdr, model


def _describe(ws, hdr, model):
    g = lambda h: txt(ws.cell(row=model, column=hdr[h]).value) if h in hdr else ""
    pk, wt = g("numberOfPackages") or "?", g("packageWeight") or "?"
    return (f"{g('serviceType')}, {pk} package{'s' if pk != '1' else ''} at {wt} "
            f"{g('weightUnits')}, from {g('senderContactName')} "
            f"({g('senderCity')} {g('senderState')})")


def _write_batch(template, out_path, people):
    wb, ws, hdr, model = _template(template)
    ncols = ws.max_column
    model_val = {c: ws.cell(row=model, column=c).value for c in range(1, ncols + 1)}
    model_fmt = {c: ws.cell(row=model, column=c).number_format for c in range(1, ncols + 1)}
    model_font = {c: copy(ws.cell(row=model, column=c).font) for c in range(1, ncols + 1)}
    by_col = {hdr[h]: f for h, f in RECIPIENT.items()}
    blank = {hdr[h] for h in BLANK_PER_RECIPIENT if h in hdr}
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)
    for i, p in enumerate(people):
        r = 2 + i
        for c in range(1, ncols + 1):
            if c in by_col:
                f = by_col[c]
                v = (phone_fmt(p[f]) if f == "Phone Number" else
                     zip5(p[f]) if f == "Zip Code" else p.get(f) or None)
            elif c in blank:
                v = None
            else:
                v = model_val[c]
            cell = ws.cell(row=r, column=c, value=v)
            cell.number_format = "@" if by_col.get(c) == "Zip Code" else model_fmt[c]
            cell.font = model_font[c]
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    wb.save(out_path)


def _read_batch_file(path):
    if not os.path.exists(path):
        sys.exit(f"File not found: {path}")
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hdr = _template_columns(ws)
    return [{f: txt(ws.cell(row=r, column=hdr[h]).value) for h, f in RECIPIENT.items()}
            for r in range(2, ws.max_row + 1)
            if txt(ws.cell(row=r, column=hdr["recipientContactName"]).value)]


# ---------------------------------------------------------------- status
def assess(cfg, cohort):
    """Everything needed to decide who can be labeled. Pure: prints nothing."""
    wb, rows, last_run = load_cohort(cfg, cohort)
    log = read_log(wb)
    current = [r for r in rows if r.get("Row State") != "Removed from report"]
    by_key = {key(r["Full Name"]): r for r in rows}
    here = [e for e in log if e.get("Cohort") == cohort]
    labeled = {key(e["Full Name"]): e for e in here}
    elsewhere = {}
    for e in log:
        if e.get("Cohort") != cohort:
            elsewhere.setdefault(key(e["Full Name"]), e)
    ready = sorted((r for r in current if is_ship_ready(cfg, r.get("Status"))),
                   key=lambda r: key(r["Full Name"]))
    fresh = [r for r in ready if key(r["Full Name"]) not in labeled]
    broken = [(r, address_problems(r)) for r in fresh if address_problems(r)]
    eligible = [r for r in fresh if not address_problems(r)]
    held = [r for r in current if not is_ship_ready(cfg, r.get("Status"))
            and key(r["Full Name"]) not in labeled]
    changed = []
    for k, e in labeled.items():
        r = by_key.get(k)
        if r is None or r.get("Row State") == "Removed from report":
            changed.append(f"{e['Full Name']}: labeled in batch {e['Batch']}, now gone from the report")
        elif not is_ship_ready(cfg, r.get("Status")):
            changed.append(f"{e['Full Name']}: labeled in batch {e['Batch']}, now '{r['Status']}'")
        else:
            moved = same_address(r, e)
            if moved:
                changed.append(f"{e['Full Name']}: {', '.join(moved).lower()} changed since "
                               f"the batch {e['Batch']} label")
    repeat = [r for r in eligible if key(r["Full Name"]) in elsewhere]
    return dict(wb=wb, log=log, last_run=last_run, ready=ready, labeled=labeled,
                batches=_batch_numbers(log, cohort), fresh=fresh, broken=broken,
                eligible=eligible, held=held, changed=changed, repeat=repeat,
                elsewhere=elsewhere)


def print_status(cfg, cohort, s):
    print(f"\n{us_date(cohort)} cohort")
    if s["last_run"]:
        hrs = (datetime.now() - s["last_run"]).total_seconds() / 3600
        age = f"{int(hrs * 60)} min ago" if hrs < 1 else f"{hrs:.0f} h ago"
        print(f"  tracker last refreshed {s['last_run']:%Y-%m-%d %H:%M} ({age})")
        if hrs > cfg.stale_hours:
            print(f"  ! That's over {cfg.stale_hours:g} hours old. Statuses may have changed "
                  f"since.\n    Refresh from Smartsheet first.")
    b = s["batches"]
    print(f"  Ready To Ship ........... {len(s['ready'])}")
    print(f"  Already labeled ......... {len(s['labeled'])}"
          + (f"   (batches {', '.join(map(str, b))})" if b else ""))
    print(f"  Need fixing first ....... {len(s['broken'])}")
    print(f"  Available now ........... {len(s['eligible'])}")
    print(f"  Not ready yet ........... {len(s['held'])}   (pending, on hold, or do-not-ship)")
    if s["changed"]:
        print("\n  ! Already labeled, but something changed (check before these ship):")
        for c in s["changed"]:
            print(f"      {c}")
    if s["broken"]:
        print("\n  Held back until fixed in Smartsheet:")
        for r, why in s["broken"]:
            print(f"      {r['Full Name']}: {'; '.join(why)}")
    if s["repeat"]:
        print("\n  ! Already got a label in an earlier cohort:")
        for r in s["repeat"]:
            e = s["elsewhere"][key(r["Full Name"])]
            print(f"      {r['Full Name']}: {us_date(e['Cohort'])} batch {e['Batch']} "
                  f"({e['Created'][:10]})")


# ---------------------------------------------------------------- operations
def make_batch(cfg, cohort, count=None, yes=False):
    s = assess(cfg, cohort)
    print_status(cfg, cohort, s)
    eligible = s["eligible"]
    if s["repeat"]:
        include = False if yes else confirm("    Include them in this batch anyway? [y/N]: ")
        if not include:
            eligible = [r for r in eligible if key(r["Full Name"]) not in s["elsewhere"]]
            print(f"    Left out{' (--yes never re-labels)' if yes else ''}. "
                  f"{len(eligible)} available.")
    if not eligible:
        print("\nNobody left to label for this cohort." if not s["fresh"] else
              "\nNobody is available until the issues above are fixed.")
        return None

    if count is not None:
        if not 1 <= count <= len(eligible):
            sys.exit(f"--count must be between 1 and {len(eligible)}.")
        n = count
    else:
        while True:
            a = ask(f"\nHow many agents in this batch? (1-{len(eligible)}, 'all', or 'q' to quit): ").lower()
            if a in ("q", "quit"):
                print("Nothing was written.")
                return None
            if a == "all":
                n = len(eligible)
                break
            if a.isdigit() and 1 <= int(a) <= len(eligible):
                n = int(a)
                break
            print(f"  Enter a number from 1 to {len(eligible)}.")
    picked = eligible[:n]

    batch = (max(s["batches"]) + 1) if s["batches"] else 1
    _, tws, thdr, tmodel = _template(cfg.template)
    print(f"\nBatch {batch} for {us_date(cohort)}: {n} agent{'s' if n != 1 else ''}")
    for i, p in enumerate(picked, 1):
        print(f"  {i:3}. {p['Full Name']:28} {p['City']}, {p['State']} {zip5(p['Zip Code'])}")
    print(f"\n  Settings from the template: {_describe(tws, thdr, tmodel)}")
    if not confirm(f"\nCreate the batch file and record these {n} as labeled? [y/N]: ", yes):
        print("Nothing was written.")
        return None

    T.assert_writable(cfg.tracker)          # lock check first: file and log never disagree
    out = os.path.join(cfg.batches, f"{cohort}_Batch-{batch}_{datetime.now():%Y-%m-%d}.xlsx")
    if os.path.exists(out):
        out = out[:-5] + f"_{datetime.now():%H%M%S}.xlsx"
    _write_batch(cfg.template, out, picked)
    try:
        _append_log(s["wb"], cohort, batch, picked, os.path.basename(out))
        T.save_tracker(cfg, s["wb"])
    except Exception as e:
        os.remove(out)
        sys.exit(f"Could not record the batch in the tracker ({e}).\n"
                 f"The batch file was removed so nothing is half-done. Fix and run again.")
    print(f"\nCreated {os.path.relpath(out, cfg.root)}")
    print(f"Recorded {n} in the Label Log. {len(eligible) - n} still to label for this cohort.")
    return out


def undo_last(cfg, cohort, yes=False):
    wb, _, _ = load_cohort(cfg, cohort)
    log = read_log(wb)
    mine = [e for e in log if e.get("Cohort") == cohort]
    if not mine:
        print(f"No batches recorded for {us_date(cohort)} yet; nothing to undo.")
        return False
    last = max(_batch_numbers(log, cohort))
    entries = [e for e in mine if e.get("Batch") == str(last)]
    print(f"Batch {last} for {us_date(cohort)}: {len(entries)} agents, created {entries[0]['Created']}")
    for e in entries:
        print(f"   {e['Full Name']}")
    if not confirm(f"\nRemove batch {last} from the Label Log so these agents can be batched "
                   f"again? [y/N]: ", yes):
        print("Nothing changed.")
        return False
    T.assert_writable(cfg.tracker)
    ws = wb[LOG_SHEET]
    for r in sorted((e["_row"] for e in entries), reverse=True):
        ws.delete_rows(r, 1)
    T.save_tracker(cfg, wb)
    print("\nRemoved. If labels were already created in FedEx, void them there too.")
    print(f"The batch file ({entries[0].get('Batch File', '?')}) was left in place.")
    return True


def import_batch(cfg, cohort, path, yes=False):
    wb, rows, _ = load_cohort(cfg, cohort)
    log = read_log(wb)
    people = _read_batch_file(path)
    if not people:
        print(f"No recipients found in {path}.")
        return False
    current = {key(r["Full Name"]): r for r in rows}
    done = {key(e["Full Name"]) for e in log if e.get("Cohort") == cohort}
    new = [p for p in people if key(p["Full Name"]) not in done]
    skip = len(people) - len(new)
    for p in new:
        p["Status"] = current.get(key(p["Full Name"]), {}).get("Status", "(not in tracker)")
    print(f"{os.path.basename(path)}: {len(people)} recipients"
          + (f", {skip} already in the log" if skip else ""))
    for p in new:
        tag = "" if key(p["Full Name"]) in current else "   <- not in the tracker for this cohort"
        print(f"   {p['Full Name']}{tag}")
    if not new:
        print("Everyone in that file is already recorded. Nothing changed.")
        return False
    if not confirm(f"\nRecord these {len(new)} as already labeled for {us_date(cohort)}? [y/N]: ", yes):
        print("Nothing changed.")
        return False
    T.assert_writable(cfg.tracker)
    nums = _batch_numbers(log, cohort)
    batch = (max(nums) + 1) if nums else 1
    _append_log(wb, cohort, batch, new, os.path.basename(path) + " (imported)")
    T.save_tracker(cfg, wb)
    print(f"\nRecorded as batch {batch}. They'll be skipped from now on.")
    return True
