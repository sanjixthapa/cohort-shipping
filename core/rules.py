"""Normalization, status classification, ordering, and address validation.
The only place these rules exist."""
import re
from datetime import datetime


# ---------------------------------------------------------------- normalization
def txt(v):
    if v is None:
        return ""
    try:                                   # pandas NaN without importing pandas here
        if v != v:
            return ""
    except Exception:
        pass
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def key(name):
    return re.sub(r"\s+", " ", txt(name)).lower()


def digits(v):
    return re.sub(r"\D", "", txt(v))


def zip5(v):
    d = re.sub(r"\D", "", txt(v).split("-")[0])
    return d.zfill(5) if d else ""


def phone10(v):
    d = digits(v)
    return d[1:] if len(d) == 11 and d.startswith("1") else d


def phone_fmt(v):
    d = phone10(v)
    return f"({d[:3]}) {d[3:6]}-{d[6:]}" if len(d) == 10 else txt(v)


def us_date(iso):
    try:
        y, m, d = iso.split("-")
        return f"{m}/{d}/{y}"
    except (ValueError, AttributeError):
        return iso


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- statuses
def is_ship_ready(cfg, status):
    return txt(status).lower() in cfg.ship_ready


def is_do_not_ship(cfg, status):
    return txt(status).lower() in cfg.do_not_ship


def sort_key(cfg, row):
    """Group by status in config order, then name. Unrecognized statuses follow the
    known pending ones; people removed from the report sit just before do-not-ship."""
    st = txt(row.get("Status")).lower()
    name = key(row.get("Full Name"))
    order = list(cfg.ship_ready) + list(cfg.pending)
    n = len(order)
    if row.get("Row State") == "Removed from report":
        return (n + 1, "", name)
    if st in order:
        return (order.index(st), "", name)
    if st in cfg.do_not_ship:
        return (n + 2 + list(cfg.do_not_ship).index(st), "", name)
    return (n, st, name)


# ---------------------------------------------------------------- validation
ADDRESS_FIELDS = ["Street Address Line 1", "Street Address Line 2", "City", "State",
                  "Zip Code", "Phone Number"]


def address_problems(p):
    """Anything that would make FedEx reject the row or misdeliver it."""
    out = []
    for f in ("Street Address Line 1", "City", "State", "Zip Code", "Phone Number"):
        if not p.get(f):
            out.append(f"no {f.lower()}")
    if p.get("State") and not re.fullmatch(r"[A-Za-z]{2}", p["State"]):
        out.append(f"state {p['State']!r}")
    if p.get("Zip Code") and len(zip5(p["Zip Code"])) != 5:
        out.append(f"zip {p['Zip Code']!r}")
    if p.get("Phone Number") and len(phone10(p["Phone Number"])) != 10:
        out.append(f"phone {p['Phone Number']!r}")
    if re.search(r"\bp\.?\s?o\.?\s?box\b",
                 f"{p.get('Street Address Line 1', '')} {p.get('Street Address Line 2', '')}", re.I):
        out.append("PO Box (FedEx can't deliver)")
    return out


def same_address(a, b):
    """Field names that differ between two address snapshots, compared normalized."""
    norm = {"Zip Code": zip5, "Phone Number": phone10}
    return [f for f in ADDRESS_FIELDS
            if norm.get(f, txt)(a.get(f, "")) != norm.get(f, txt)(b.get(f, ""))]


# ---------------------------------------------------------------- cohort detection
def pick_cohort(dates, today=None):
    """Earliest cohort on or after today; if none, the latest one. dates are ISO strings."""
    dates = sorted(set(d for d in dates if d))
    if not dates:
        return None
    today = (today or datetime.now()).strftime("%Y-%m-%d")
    upcoming = [d for d in dates if d >= today]
    return upcoming[0] if upcoming else dates[-1]
