"""Oracle Open Item Report loader and in-memory indexes."""
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from .keys import clean
from .references import norm_ref

HEADERS = {
    "trx": "Trx Number",
    "date": "Trx Date",
    "due": "Due Remaining",
    "party": "Party Name",
    "bill_to": "Bill To Cust",
    "ship_to": "Ship To Loc",
    "store": "Store",
}


def to_date(value):
    if value is None or (isinstance(value, float) and value != value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%Y%m%d", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


class OpenItems:
    """Open items as a list of dict rows plus lookup indexes."""

    def __init__(self, rows, key_func, store_shiptos=None):
        self.rows = rows
        self.key_func = key_func
        self.store_shiptos = store_shiptos or {}
        self.by_key = defaultdict(list)
        self.by_shipto_date = defaultdict(list)
        self.by_id = {}
        self.refs = set()
        for r in rows:
            self.by_id[r["id"]] = r
            self.by_key[r["key"]].append(r)
            if r["ship_to"] and r["date"]:
                self.by_shipto_date[(r["ship_to"], r["date"])].append(r)
            self.refs.add(norm_ref(r["trx"]))
        # store -> bill-tos, via the ship-tos of that store
        self.store_billtos = defaultdict(set)
        shipto_billto = defaultdict(set)
        for r in rows:
            if r["ship_to"] and r["bill_to"]:
                shipto_billto[r["ship_to"]].add(r["bill_to"])
        for store, shiptos in self.store_shiptos.items():
            for s in shiptos:
                self.store_billtos[store] |= shipto_billto.get(s, set())

    def subset(self, bill_tos):
        keep = set(bill_tos)
        return OpenItems([r for r in self.rows if r["bill_to"] in keep], self.key_func, self.store_shiptos)

    def __len__(self):
        return len(self.rows)


def _clean_id(v):
    s = clean(v)
    return s if s and s.lower() != "nan" else ""


# Demo API: each report is read from Excel once and kept in memory. Every run gets its own copy of the rows,
# so one request cannot affect the next. A changed file (new modification time) is read again.
_CACHE = {}


def load_open_items(path, key_func):
    p = Path(path).resolve()
    stamp = (str(p), p.stat().st_mtime, key_func)
    if stamp not in _CACHE:
        for old in [k for k in _CACHE if k[0] == str(p)]:
            del _CACHE[old]
        _CACHE[stamp] = _read_open_items(p, key_func)
    rows, store_shiptos = _CACHE[stamp]
    return OpenItems([dict(r) for r in rows], key_func, {s: set(v) for s, v in store_shiptos.items()})


def cached_reports():
    """{file name: rows} of the reports currently held in memory."""
    return {Path(k[0]).name: len(v[0]) for k, v in _CACHE.items()}


def _read_open_items(path, key_func):
    df = pd.read_excel(path, dtype=object)
    rows = []
    store_shiptos = defaultdict(set)
    cols = {k: (h if h in df.columns else None) for k, h in HEADERS.items()}
    if cols["trx"] is None:
        raise ValueError(f"'{HEADERS['trx']}' column not found in {path}")
    for i, rec in enumerate(df.to_dict("records")):
        trx = clean(rec.get(cols["trx"]))
        if not trx:
            continue
        due = pd.to_numeric(rec.get(cols["due"]), errors="coerce") if cols["due"] else 0.0
        row = {
            "id": i,
            "trx": trx,
            "key": key_func(trx),
            "date": to_date(rec.get(cols["date"])) if cols["date"] else None,
            "due": round(float(due), 2) if due == due else 0.0,
            "party": clean(rec.get(cols["party"])) if cols["party"] else "",
            "bill_to": _clean_id(rec.get(cols["bill_to"])) if cols["bill_to"] else "",
            "ship_to": _clean_id(rec.get(cols["ship_to"])) if cols["ship_to"] else "",
            "store": _clean_id(rec.get(cols["store"])) if cols["store"] else "",
        }
        rows.append(row)
        if row["store"] and row["ship_to"]:
            store_shiptos[row["store"]].add(row["ship_to"])
    # Costco report carries a separate Customer -> Store No. lookup table
    if "Customer" in df.columns and "Store No." in df.columns:
        for cust, store in df[["Customer", "Store No."]].dropna().itertuples(index=False):
            store_shiptos[_clean_id(store)].add(_clean_id(cust))
    return rows, dict(store_shiptos)
