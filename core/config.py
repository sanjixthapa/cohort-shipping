"""Loads config.toml once. Every setting the tools use comes from here."""
import os
import sys
import tomllib
from dataclasses import dataclass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE = os.path.join(ROOT, "config.toml")

DEFAULTS = {
    "smartsheet": {"report_id": None},
    "files": {"tracker": "Cohort_Tracker.xlsx",
              "template": "Call_center_batch_template.xlsx",
              "batches_folder": "batches"},
    "statuses": {"ship_ready": ["Ready To Ship"],
                 "pending": ["Ready to start", "Pending Background check",
                             "Pending offer stage", "Pushed Pending"],
                 "do_not_ship": ["Equipment Shipped", "Started", "Dropped",
                                 "Terminated", "Did Not Start"]},
    "checks": {"stale_hours": 12},
}


@dataclass(frozen=True)
class Config:
    root: str
    report_id: int | None      # None until set in config.toml
    tracker: str
    template: str
    batches: str
    token_file: str
    ship_ready: tuple          # lowercase
    pending: tuple             # lowercase, display order
    do_not_ship: tuple         # lowercase, display order
    stale_hours: float


def _abs(p):
    return p if os.path.isabs(p) else os.path.join(ROOT, p)


def load():
    data = {k: dict(v) for k, v in DEFAULTS.items()}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "rb") as f:
                user = tomllib.load(f)
        except tomllib.TOMLDecodeError as e:
            sys.exit(f"config.toml has a typo: {e}\nFix it, or delete it to use the defaults.")
        for section, values in user.items():
            if section in data and isinstance(values, dict):
                data[section].update(values)
    st = data["statuses"]
    low = lambda xs: tuple(str(x).strip().lower() for x in xs)
    ship, pend, dns = low(st["ship_ready"]), low(st["pending"]), low(st["do_not_ship"])
    overlap = (set(ship) & set(pend)) | (set(ship) & set(dns)) | (set(pend) & set(dns))
    if overlap:
        sys.exit(f"config.toml lists these statuses in more than one group: {sorted(overlap)}")
    if not ship:
        sys.exit("config.toml: statuses.ship_ready can't be empty.")
    rid = data["smartsheet"].get("report_id")
    if rid in (None, "", 0):
        rid = None
    else:
        try:
            rid = int(rid)
        except (TypeError, ValueError):
            sys.exit("config.toml: smartsheet.report_id must be a number.")
    return Config(
        root=ROOT, report_id=rid,
        tracker=_abs(data["files"]["tracker"]),
        template=_abs(data["files"]["template"]),
        batches=_abs(data["files"]["batches_folder"]),
        token_file=os.path.join(ROOT, "smartsheet_token.txt"),
        ship_ready=ship, pending=pend, do_not_ship=dns,
        stale_hours=float(data["checks"]["stale_hours"]),
    )
