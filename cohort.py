#!/usr/bin/env python3
"""
Cohort shipping: one command for the whole routine.

    py cohort.py                     menu (or double-click run.bat)

    py cohort.py refresh             pull Smartsheet into the tracker
    py cohort.py status              where the cohort stands
    py cohort.py labels              refresh, then build a FedEx batch (asks how many)
    py cohort.py undo                take the last batch back out of the Label Log
    py cohort.py import FILE         record a batch you built by hand

The cohort is picked automatically: the earliest one on or after today.
Override with --cohort YYYY-MM-DD. Settings live in config.toml.
"""
import argparse
import os
import subprocess
import sys
from datetime import datetime

from core import config, labels, source, tracker
from core.rules import pick_cohort, us_date
from core.ui import ask, confirm


# ---------------------------------------------------------------- cohort selection
def valid_date(s):
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return s
    except ValueError:
        raise argparse.ArgumentTypeError(f"expected YYYY-MM-DD, got {s!r}")


def tracker_cohorts(cfg):
    wb = tracker.open_tracker(cfg)
    return tracker.cohort_sheets(wb)


def cohort_from_tracker(cfg, given=None):
    if given:
        return given
    c = pick_cohort(tracker_cohorts(cfg))
    if c is None:
        sys.exit("The tracker has no cohorts yet. Refresh from Smartsheet first "
                 "(menu option 1, or: py cohort.py refresh).")
    return c


# ---------------------------------------------------------------- operations
def do_refresh(cfg, cohort=None, path=None):
    df = source.load_report(cfg, path)
    if cohort is None:
        cohort = pick_cohort(source.cohort_dates(df))
        if cohort is None:
            sys.exit("The report has no cohort dates.")
        print(f"Cohort: {us_date(cohort)} (earliest on or after today)")
    tracker.refresh(cfg, df, cohort)
    return cohort


def do_labels(cfg, cohort=None, refresh=True, count=None, yes=False, path=None):
    if refresh:
        cohort = do_refresh(cfg, cohort, path)
        print()
    else:
        cohort = cohort_from_tracker(cfg, cohort)
    labels.make_batch(cfg, cohort, count=count, yes=yes)
    return cohort


def open_path(p):
    if not os.path.exists(p):
        print(f"{p} doesn't exist yet.")
        return
    if hasattr(os, "startfile"):
        os.startfile(p)                          # Windows: opens with the default app
    else:
        subprocess.run(["xdg-open" if sys.platform != "darwin" else "open", p], check=False)


# ---------------------------------------------------------------- menu
MENU = """
  1  Refresh from Smartsheet
  2  Show status
  3  Make a label batch        (refreshes first)
  4  Undo the last batch
  5  Record a hand-made batch
  6  Open the tracker in Excel
  7  Open the batches folder
  C  Change cohort
  Q  Quit
"""


def headline(cfg, cohort):
    if cohort is None:
        return "Cohort: picked automatically on first refresh"
    try:
        s = labels.assess(cfg, cohort)
    except SystemExit:
        return f"Cohort: {us_date(cohort)}   (not in the tracker yet; refresh first)"
    when = f"refreshed {s['last_run']:%m/%d %H:%M}" if s["last_run"] else "not refreshed"
    return (f"Cohort: {us_date(cohort)}   {when}\n"
            f"Ready {len(s['ready'])}  ·  labeled {len(s['labeled'])}  ·  "
            f"available {len(s['eligible'])}  ·  need fixing {len(s['broken'])}")


def change_cohort(cfg, current):
    have = tracker_cohorts(cfg)
    if have:
        print("\nCohorts in the tracker:")
        for i, c in enumerate(have, 1):
            print(f"  {i}  {us_date(c)}{'   (current)' if c == current else ''}")
    a = ask("\nPick a number, type a date (YYYY-MM-DD) for a new cohort, or Enter for automatic: ")
    if not a:
        return None, False
    if a.isdigit() and 1 <= int(a) <= len(have):
        return have[int(a) - 1], True
    try:
        return valid_date(a), True
    except argparse.ArgumentTypeError:
        print("  Not a number from the list or a YYYY-MM-DD date. Cohort unchanged.")
        return current, current is not None


def menu(cfg):
    pinned = False                                  # True once you pick a cohort yourself
    try:
        cohort = pick_cohort(tracker_cohorts(cfg))
    except SystemExit:
        cohort = None
    while True:
        print("\n" + "=" * 60)
        print(" Cohort shipping")
        print(" " + headline(cfg, cohort).replace("\n", "\n "))
        print("=" * 60 + MENU)
        choice = ask("Choose: ").lower()
        try:
            if choice in ("q", "quit", "exit"):
                return
            elif choice == "1":
                cohort = do_refresh(cfg, cohort if pinned else None)
            elif choice == "2":
                c = cohort_from_tracker(cfg, cohort)
                labels.print_status(cfg, c, labels.assess(cfg, c))
            elif choice == "3":
                try:
                    cohort = do_refresh(cfg, cohort if pinned else None)
                    print()
                except SystemExit as e:
                    print(f"\nRefresh failed: {e}")
                    if not confirm("Make the batch from the tracker as it is? [y/N]: "):
                        continue
                labels.make_batch(cfg, cohort_from_tracker(cfg, cohort))
            elif choice == "4":
                labels.undo_last(cfg, cohort_from_tracker(cfg, cohort))
            elif choice == "5":
                p = ask("Path to the batch file (you can drag the file into this window): ")
                p = p.strip().strip('"').strip("'")
                if p:
                    labels.import_batch(cfg, cohort_from_tracker(cfg, cohort), p)
            elif choice == "6":
                open_path(cfg.tracker)
                print("Close it in Excel before refreshing or making labels.")
                continue
            elif choice == "7":
                os.makedirs(cfg.batches, exist_ok=True)
                open_path(cfg.batches)
                continue
            elif choice == "c":
                cohort, pinned = change_cohort(cfg, cohort)
                if cohort is None:
                    try:
                        cohort = pick_cohort(tracker_cohorts(cfg))
                    except SystemExit:
                        cohort = None
                continue
            else:
                print("  Pick one of the options above.")
                continue
        except SystemExit as e:                     # an error message, not a crash
            if e.code not in (None, 0):
                print(f"\n{e}")
        ask("\nPress Enter to return to the menu...")


# ---------------------------------------------------------------- command line
def main():
    cfg = config.load()
    if len(sys.argv) == 1:
        menu(cfg)
        return
    ap = argparse.ArgumentParser(prog="cohort.py", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--cohort", type=valid_date, help="YYYY-MM-DD (default: automatic)")

    p = sub.add_parser("refresh", parents=[common], help="pull Smartsheet into the tracker")
    p.add_argument("--file", help="use a local .xlsx export instead of Smartsheet")
    sub.add_parser("status", parents=[common], help="where the cohort stands")
    p = sub.add_parser("labels", parents=[common], help="build a FedEx batch")
    p.add_argument("--count", type=int, help="skip the question and take this many")
    p.add_argument("--yes", action="store_true", help="skip the final confirmation")
    p.add_argument("--no-refresh", action="store_true", help="use the tracker as it is")
    p.add_argument("--file", help="refresh from a local .xlsx export instead of Smartsheet")
    p = sub.add_parser("undo", parents=[common], help="take the last batch out of the log")
    p.add_argument("--yes", action="store_true")
    p = sub.add_parser("import", parents=[common], help="record a hand-made batch")
    p.add_argument("file")
    p.add_argument("--yes", action="store_true")
    a = ap.parse_args()

    if a.cmd == "refresh":
        do_refresh(cfg, a.cohort, a.file)
    elif a.cmd == "status":
        c = cohort_from_tracker(cfg, a.cohort)
        labels.print_status(cfg, c, labels.assess(cfg, c))
    elif a.cmd == "labels":
        do_labels(cfg, a.cohort, refresh=not a.no_refresh, count=a.count, yes=a.yes, path=a.file)
    elif a.cmd == "undo":
        labels.undo_last(cfg, cohort_from_tracker(cfg, a.cohort), yes=a.yes)
    elif a.cmd == "import":
        labels.import_batch(cfg, cohort_from_tracker(cfg, a.cohort), a.file, yes=a.yes)


if __name__ == "__main__":
    main()
