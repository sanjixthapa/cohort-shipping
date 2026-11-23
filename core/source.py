"""Where the onboarding report comes from: live from Smartsheet, or a local .xlsx export.
Both return the same DataFrame. Read-only: only GET requests. Standard library only."""
import getpass
import json
import os
import sys
import time
import urllib.error
import urllib.request

import pandas as pd

API = "https://api.smartsheet.com/2.0"
TOKEN_ENV = "SMARTSHEET_API_TOKEN"
PAGE_SIZE = 500
MAX_PAGES = 200
RAW_VALUE_TYPES = {"DATE", "DATETIME", "ABSTRACT_DATETIME", "CHECKBOX"}
_token_source = None


# ---------------------------------------------------------------- token
def get_token(cfg):
    """Environment variable, then smartsheet_token.txt, then ask (and offer to save)."""
    global _token_source
    tok = os.environ.get(TOKEN_ENV, "").strip()
    if tok:
        _token_source = f"the {TOKEN_ENV} environment variable"
        return tok
    if os.path.exists(cfg.token_file):
        with open(cfg.token_file, encoding="utf-8-sig") as f:
            tok = f.read().strip()
        if tok:
            _token_source = cfg.token_file
            return tok
    if sys.stdin.isatty():
        tok = getpass.getpass("Paste your Smartsheet API token (input is hidden): ").strip()
        if tok:
            _token_source = "what you typed"
            if input("Save it so you aren't asked again? [y/N]: ").strip().lower() in ("y", "yes"):
                with open(cfg.token_file, "w", encoding="utf-8") as f:
                    f.write(tok + "\n")
                print(f"  Saved to {cfg.token_file}")
                print("  Don't share or upload that file. Delete it to stop using the saved token.")
            return tok
    sys.exit("No Smartsheet API token available.\n"
             f"Run from a terminal so it can ask you, or put the token in {cfg.token_file}.")


# ---------------------------------------------------------------- http
def _explain(cfg, code, msg, report_id):
    if code == 401:
        fix = ""
        if _token_source == cfg.token_file:
            fix = (f"\nThe token came from {cfg.token_file}. Delete that file and run again; "
                   f"you'll be asked for the new token and can save it.")
        elif _token_source:
            fix = f"\nThe token came from {_token_source}."
        return ("Smartsheet rejected the token (401). It may be mistyped, revoked, or expired.\n"
                "Generate a new one: profile icon > Personal Settings > API Access." + fix)
    if code == 403:
        return (f"Access denied (403) for report {report_id}.\nEither your account can't open "
                f"that report, or API access is switched off for your account (an admin controls that).")
    if code == 404:
        return (f"Report {report_id} was not found (404). Check smartsheet.report_id in "
                f"config.toml, and that the report hasn't been moved or deleted.")
    return f"Smartsheet returned HTTP {code}: {msg}"


def _get_json(cfg, url, token, report_id, retries=4):
    for attempt in range(retries):
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}",
                                                   "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            raw = e.read().decode("utf-8", "replace") if hasattr(e, "read") else ""
            try:
                msg = json.loads(raw).get("message", raw)
            except Exception:
                msg = raw
            if e.code in (429, 500, 502, 503, 504) and attempt < retries - 1:
                wait = 2 ** (attempt + 1)
                print(f"  Smartsheet said {e.code}; retrying in {wait}s...")
                time.sleep(wait)
                continue
            sys.exit(_explain(cfg, e.code, msg, report_id))
        except urllib.error.URLError as e:
            sys.exit(f"Could not reach Smartsheet: {e.reason}\nIf this says certificate verify "
                     f"failed, your network is inspecting HTTPS traffic; ask IT for the "
                     f"proxy/certificate settings.")
    sys.exit("Smartsheet did not respond after several attempts.")


# ---------------------------------------------------------------- fetch
def fetch_report_df(cfg, report_id, token):
    rows, cols, total, page = [], None, 0, 1
    while page <= MAX_PAGES:
        d = _get_json(cfg, f"{API}/reports/{report_id}?pageSize={PAGE_SIZE}&page={page}",
                      token, report_id)
        if cols is None:
            cols = {c["virtualId"]: (c["title"], c.get("type", "")) for c in d.get("columns", [])}
            total = d.get("totalRowCount", 0)
            if not cols:
                sys.exit(f"Report {report_id} returned no columns.")
        batch = d.get("rows", [])
        rows.extend(batch)
        if not batch or len(rows) >= total:
            break
        page += 1
    recs = []
    for r in rows:
        rec = {}
        for c in r.get("cells", []):
            title, typ = cols.get(c.get("virtualColumnId"), (None, None))
            if title is None:
                continue
            if typ in RAW_VALUE_TYPES:
                v = c.get("value")
            else:
                v = c.get("displayValue")
                if v in (None, ""):
                    v = c.get("value")
            rec[title] = v
        recs.append(rec)
    df = pd.DataFrame(recs, columns=[t for t, _ in cols.values()])
    if "Cohort Date" in df.columns:
        df["Cohort Date"] = pd.to_datetime(df["Cohort Date"], errors="coerce")
    if len(df) != total:
        print(f"  warning: Smartsheet reports {total} rows but {len(df)} were received")
    return df


def load_report(cfg, path=None):
    if path:
        try:
            return pd.read_excel(path, header=0)
        except FileNotFoundError:
            sys.exit(f"Report not found: {path}")
        except Exception as e:
            sys.exit(f"Could not read {path}: {e}")
    if cfg.report_id is None:
        sys.exit("No Smartsheet report is set. Add report_id under [smartsheet] in config.toml "
                 "(see config.example.toml), or use --file with a local .xlsx export.")
    df = fetch_report_df(cfg, cfg.report_id, get_token(cfg))
    print(f"Fetched {len(df)} rows live from Smartsheet report {cfg.report_id}")
    return df


def cohort_dates(df):
    return sorted({str(d)[:10] for d in df["Cohort Date"].dropna()}) if "Cohort Date" in df else []
