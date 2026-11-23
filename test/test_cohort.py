"""Tests for the cohort shipping tool. Standard library only:

    python -m unittest discover -s tests -v

Everything runs on the made-up data in sample/, in a throwaway folder.
"""
import builtins
import contextlib
import io
import os
import shutil
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import openpyxl                                   # noqa: E402
import pandas as pd                               # noqa: E402

from core import config, labels, rules, source, tracker   # noqa: E402
from core.config import Config                    # noqa: E402

SAMPLE_REPORT = os.path.join(ROOT, "sample", "sample_report.xlsx")
SAMPLE_TEMPLATE = os.path.join(ROOT, "sample", "sample_fedex_template.xlsx")
COHORT_A, COHORT_B = "2027-01-07", "2027-01-28"


def quiet(fn, *a, **k):
    with contextlib.redirect_stdout(io.StringIO()):
        return fn(*a, **k)


class Workspace(unittest.TestCase):
    """Each test gets a fresh folder with the sample template and a config."""

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        shutil.copy(SAMPLE_TEMPLATE, os.path.join(self.dir, "template.xlsx"))
        self.cfg = Config(
            root=self.dir, report_id=None,
            tracker=os.path.join(self.dir, "Cohort_Tracker.xlsx"),
            template=os.path.join(self.dir, "template.xlsx"),
            batches=os.path.join(self.dir, "batches"),
            token_file=os.path.join(self.dir, "smartsheet_token.txt"),
            ship_ready=("ready to ship",),
            pending=("ready to start", "pending background check",
                     "pending offer stage", "pushed pending"),
            do_not_ship=("equipment shipped", "started", "dropped", "terminated",
                         "did not start"),
            stale_hours=12)
        self.df = pd.read_excel(SAMPLE_REPORT)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def refresh(self, df=None, cohort=COHORT_A):
        return quiet(tracker.refresh, self.cfg, self.df if df is None else df, cohort)

    def sheet(self, cohort=COHORT_A):
        ws = openpyxl.load_workbook(self.cfg.tracker)[cohort]
        hr = tracker.header_row(ws)
        cols = {ws.cell(row=hr, column=c).value: c for c in range(1, ws.max_column + 1)}
        return ws, cols, {ws.cell(row=r, column=1).value: r
                          for r in range(hr + 1, ws.max_row + 1) if ws.cell(row=r, column=1).value}

    def batch_names(self, path):
        ws = openpyxl.load_workbook(path)["Overview"]
        return [ws.cell(row=r, column=18).value for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=18).value]


# ---------------------------------------------------------------- rules
class RulesTests(unittest.TestCase):
    def test_zip_keeps_leading_zero(self):
        self.assertEqual(rules.zip5(2108), "02108")
        self.assertEqual(rules.zip5(2108.0), "02108")
        self.assertEqual(rules.zip5("02108-1234"), "02108")
        self.assertEqual(rules.zip5(None), "")

    def test_phone_drops_country_code(self):
        self.assertEqual(rules.phone_fmt(15125550101), "(512) 555-0101")
        self.assertEqual(rules.phone_fmt(15125550101.0), "(512) 555-0101")
        self.assertEqual(rules.phone10("(512) 555-0101"), "5125550101")

    def test_pick_cohort(self):
        from datetime import datetime
        dates = ["2026-11-19", "2026-12-10", "2027-01-01"]
        self.assertEqual(rules.pick_cohort(dates, datetime(2026, 11, 20)), "2026-12-10")
        self.assertEqual(rules.pick_cohort(dates, datetime(2026, 12, 10)), "2026-12-10")
        self.assertEqual(rules.pick_cohort(dates, datetime(2027, 6, 1)), "2027-01-01")
        self.assertIsNone(rules.pick_cohort([]))

    def test_address_problems(self):
        good = {"Street Address Line 1": "1 Main St", "City": "Austin", "State": "TX",
                "Zip Code": "78701", "Phone Number": "(512) 555-0101"}
        self.assertEqual(rules.address_problems(good), [])
        self.assertTrue(rules.address_problems({**good, "State": "TX`"}))
        self.assertTrue(rules.address_problems({**good, "Street Address Line 1": "PO Box 12"}))
        self.assertTrue(rules.address_problems({**good, "City": ""}))
        self.assertTrue(rules.address_problems({**good, "Phone Number": "555-0101"}))


# ---------------------------------------------------------------- config
class ConfigTests(unittest.TestCase):
    def load_with(self, text):
        d = tempfile.mkdtemp()
        path = os.path.join(d, "config.toml")
        if text is not None:
            with open(path, "w") as f:
                f.write(text)
        old = config.CONFIG_FILE
        config.CONFIG_FILE = path
        try:
            return config.load()
        finally:
            config.CONFIG_FILE = old
            shutil.rmtree(d, ignore_errors=True)

    def test_defaults_without_a_file(self):
        cfg = self.load_with(None)
        self.assertIsNone(cfg.report_id)
        self.assertIn("ready to ship", cfg.ship_ready)

    def test_report_id_from_file(self):
        self.assertEqual(self.load_with("[smartsheet]\nreport_id = 42\n").report_id, 42)

    def test_typo_is_reported(self):
        with self.assertRaises(SystemExit):
            self.load_with("[statuses\n")

    def test_status_in_two_groups_is_rejected(self):
        with self.assertRaises(SystemExit):
            self.load_with('[statuses]\nship_ready=["A"]\npending=["A"]\n')

    def test_no_report_id_means_no_live_fetch(self):
        cfg = self.load_with(None)
        with self.assertRaises(SystemExit) as e:
            source.load_report(cfg)
        self.assertIn("report_id", str(e.exception))


# ---------------------------------------------------------------- tracker
class TrackerTests(Workspace):
    def test_first_refresh_adds_everyone(self):
        log = self.refresh()
        expected = int((self.df["Cohort Date"] == pd.Timestamp(COHORT_A)).sum())
        self.assertEqual(sum(1 for e in log if e[3] == "Added"), expected)

    def test_second_refresh_changes_nothing(self):
        self.refresh()
        self.assertEqual(self.refresh(), [])

    def test_status_change_is_updated_and_logged(self):
        self.refresh()
        df = self.df.copy()
        df.loc[df["Full Name"] == "Khan, Zara", "Status"] = "Ready To Ship"
        log = self.refresh(df)
        self.assertEqual([(e[2], e[4], e[6]) for e in log],
                         [("Khan, Zara", "Status", "Ready To Ship")])
        ws, cols, rows = self.sheet()
        self.assertEqual(ws.cell(row=rows["Khan, Zara"], column=cols["Row State"]).value, "Updated")

    def test_removed_person_is_kept_and_flagged(self):
        self.refresh()
        self.refresh(self.df[self.df["Full Name"] != "Ito, Kenji"])
        ws, cols, rows = self.sheet()
        self.assertIn("Ito, Kenji", rows)
        self.assertEqual(ws.cell(row=rows["Ito, Kenji"], column=cols["Row State"]).value,
                         "Removed from report")

    def test_your_columns_are_never_overwritten(self):
        self.refresh()
        wb = openpyxl.load_workbook(self.cfg.tracker)
        ws = wb[COHORT_A]
        _, cols, rows = self.sheet()
        ws.cell(row=rows["Jones, Ava"], column=cols["Tracking Number"], value="7712 0000 1111")
        wb.save(self.cfg.tracker)
        df = self.df.copy()
        df.loc[df["Full Name"] == "Jones, Ava", "City"] = "Decatur"   # force an update
        self.refresh(df)
        ws, cols, rows = self.sheet()
        self.assertEqual(ws.cell(row=rows["Jones, Ava"], column=cols["Tracking Number"]).value,
                         "7712 0000 1111")

    def test_leading_zero_zip_in_tracker(self):
        self.refresh()
        ws, cols, rows = self.sheet()
        self.assertEqual(ws.cell(row=rows["Chen, Lily"], column=cols["Zip Code"]).value, "02108")


# ---------------------------------------------------------------- labels
class LabelTests(Workspace):
    def setUp(self):
        super().setUp()
        self.refresh()

    def make(self, n, cohort=COHORT_A):
        return quiet(labels.make_batch, self.cfg, cohort, count=n, yes=True)

    def test_batches_never_overlap_and_cover_everyone(self):
        s = labels.assess(self.cfg, COHORT_A)
        eligible = [r["Full Name"] for r in s["eligible"]]
        names = []
        for n in (4, 3, len(eligible) - 7):
            names += self.batch_names(self.make(n))
        self.assertEqual(len(names), len(set(names)), "someone was labeled twice")
        self.assertEqual(names, eligible, "batches should continue alphabetically")
        self.assertEqual(labels.assess(self.cfg, COHORT_A)["eligible"], [])
        self.assertIsNone(self.make(1), "a fourth batch should have nobody to label")

    def test_bad_addresses_and_unknown_statuses_are_held_back(self):
        s = labels.assess(self.cfg, COHORT_A)
        held = {r["Full Name"] for r, _ in s["broken"]}
        self.assertEqual(held, {"Garcia, Elena", "Hughes, Owen"})       # PO Box, "TX`"
        eligible = {r["Full Name"] for r in s["eligible"]}
        self.assertNotIn("Patel, Ravi", eligible)                       # "On Hold"
        self.assertNotIn("Khan, Zara", eligible)                        # pending

    def test_batch_file_values(self):
        out = self.make(len(labels.assess(self.cfg, COHORT_A)["eligible"]))
        ws = openpyxl.load_workbook(out)["Overview"]
        row = {ws.cell(row=r, column=18).value: r for r in range(2, ws.max_row + 1)}
        self.assertEqual(ws.cell(row=row["Chen, Lily"], column=23).value, "02108")
        self.assertEqual(ws.cell(row=row["Okafor, Chidi"], column=23).value, "07102")
        self.assertEqual(ws.cell(row=row["Chen, Lily"], column=20).value, "(617) 555-0103")
        tpl = openpyxl.load_workbook(self.cfg.template)["Overview"]
        for r in row.values():                   # sender block copied from the template
            for c in range(1, 18):
                self.assertEqual(ws.cell(row=r, column=c).value, tpl.cell(row=2, column=c).value)
        self.assertNotIn("Sample, Recipient", row)                      # old rows cleared

    def test_undo_brings_people_back(self):
        before = len(labels.assess(self.cfg, COHORT_A)["eligible"])
        self.make(3)
        quiet(labels.undo_last, self.cfg, COHORT_A, yes=True)
        self.assertEqual(len(labels.assess(self.cfg, COHORT_A)["eligible"]), before)

    def test_redo_after_undo_never_overwrites_a_file(self):
        first = self.make(2)
        quiet(labels.undo_last, self.cfg, COHORT_A, yes=True)
        second = self.make(2)
        self.assertNotEqual(first, second)
        self.assertTrue(os.path.exists(first) and os.path.exists(second))

    def test_label_from_an_earlier_cohort_is_not_repeated(self):
        self.make(1)                                                    # labels Alvarez in A
        df = self.df.copy()
        df.loc[df["Full Name"] == "Alvarez, Maya", "Cohort Date"] = pd.Timestamp(COHORT_B)
        self.refresh(df, COHORT_B)
        s = labels.assess(self.cfg, COHORT_B)
        self.assertIn("Alvarez, Maya", [r["Full Name"] for r in s["repeat"]])
        out = self.make(1, COHORT_B)                                    # --yes never re-labels
        self.assertNotIn("Alvarez, Maya", self.batch_names(out))

    def test_nothing_written_while_excel_has_the_tracker_open(self):
        real_open = builtins.open

        def locked(path, mode="r", *a, **k):
            if os.path.abspath(str(path)) == self.cfg.tracker and "+" in mode:
                raise PermissionError(13, "being used by another process")
            return real_open(path, mode, *a, **k)

        builtins.open = locked
        try:
            with self.assertRaises(SystemExit):
                self.make(2)
        finally:
            builtins.open = real_open
        self.assertFalse(os.path.exists(self.cfg.batches) and os.listdir(self.cfg.batches))
        self.assertEqual(labels.read_log(openpyxl.load_workbook(self.cfg.tracker)), [])


if __name__ == "__main__":
    unittest.main()
