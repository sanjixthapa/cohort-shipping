# Cohort Shipping Automation

Automates equipment shipping for new-hire cohorts: pulls the live roster from a
Smartsheet report, tracks changes as the cohort fills in, and builds FedEx batch
label files a few agents at a time, without ever labeling anyone twice.

## The problem

Every three weeks a cohort of 50 to 70 remote new hires needs equipment shipped
before their start date. The roster lives in Smartsheet and keeps changing until
the last minute: people are added, clear background checks, drop out, or move to a
later cohort. Labels go out over three or four days.

Doing this by hand meant exporting the report, filtering by status, copying addresses
into FedEx's batch template, and remembering who had already been labeled. That
process produced real problems: people already shipped in an earlier cohort showing
up again, ZIP codes losing their leading zero in Excel, and data-entry errors (a
stray backtick in a state code, a PO-box-only ZIP) that a carrier can't deliver to.

## What it does

- **Refreshes a tracker workbook** from Smartsheet. New people are appended, changes
  are updated and logged with before/after values, and people who leave the report
  are kept and flagged rather than silently deleted.
- **Builds FedEx batch files on demand.** It asks how many agents to include, shows
  exactly who that is, and only writes after confirmation.
- **Guarantees no duplicate labels.** Every batch is recorded in a Label Log, and
  later batches pick up where the last one stopped.
- **Catches problems before they ship:**
  - **Labeled earlier, now different:** someone was labeled, then dropped or changed address.
  - **Repeat recipients:** someone already labeled in an earlier cohort.
  - **Bad address data:** invalid state codes, short ZIPs, bad phone numbers, missing fields, PO boxes.
  - **Unrecognized statuses:** these are never treated as ready to ship.

## Design decisions

- **One source of rules.** Status vocabulary, sorting, normalization, and address
  validation live in `core/rules.py` and `config.toml`. New statuses are a config
  change, not a code change.
- **Read-only integration.** The Smartsheet client only issues GET requests, uses
  only the Python standard library, and retries on rate limits and transient errors.
- **Fail before writing, never halfway.** If the tracker is open in Excel, the tool
  stops before creating anything. If recording a batch fails, the batch file is
  removed. The batch file and the Label Log never disagree.
- **Never destructive.** No row is ever deleted from the tracker, batch files are
  never overwritten, and a backup is written before every save.
- **Users own their columns.** Manual columns (shipped?, tracking number, notes) are
  preserved across every refresh.

## Architecture

```
Smartsheet report ──► core/source.py ──► core/tracker.py ──► Cohort_Tracker.xlsx
   (read-only)           (or .xlsx)        merge + change log     cohort sheets
                                                                  Change Log
                                                                  Label Log
                                                                      │
FedEx template ─────────────────────────► core/labels.py ◄────────────┘
                                                 │
                                                 ▼
                                    batches/<cohort>_Batch-N.xlsx

cohort.py: menu and command line        core/rules.py: shared rules
core/config.py: reads config.toml       core/ui.py: prompts
```

## Try it with the sample data

The `sample/` folder contains a made-up roster and FedEx template. No Smartsheet
account needed.

```
pip install -r requirements.txt
cp sample/sample_fedex_template.xlsx Call_center_batch_template.xlsx

python cohort.py refresh --file sample/sample_report.xlsx
python cohort.py status
python cohort.py labels --no-refresh
```

The sample roster deliberately includes a PO box, a state code with a stray
backtick, ZIP codes that lost their leading zero, and an unrecognized status, so the
safeguards are visible.

On Windows, `run.bat` opens an interactive menu.

## Using it with Smartsheet

1. Copy `config.example.toml` to `config.toml` and set `report_id`.
2. Generate an API token in Smartsheet: Personal Settings, then API Access.
3. Run `python cohort.py`. You're asked for the token once, and can save it locally.

`config.toml`, the token, and every spreadsheet except the sample data are
git-ignored.

## Tests

```
python -m unittest discover -s tests -v
```

22 tests run on the sample data in a temporary folder. They cover normalization,
config validation, tracker merging, and the labeling safeguards: no overlap across
batches, held-back addresses, undo, cross-cohort repeats, and refusing to write while
the tracker is locked.

## Requirements

Python 3.11 or newer, `pandas`, `openpyxl`.
