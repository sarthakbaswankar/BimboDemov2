"""Metro mixed-record remittance CSV parser.

Column 1 is the record type. The first row of each record type is its header
definition (e.g. 'REMITTANCE,Control_Number,Store Number,...'); data rows follow.
"""
import csv
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd


@dataclass
class MetroRemittance:
    path: Path
    region_code: str | None
    file_date: object
    summaries: list = field(default_factory=list)   # dicts: control, vendor, name, total
    lines: pd.DataFrame = None                       # REMITTANCE rows
    other: pd.DataFrame = None                       # DSD ADJ / MARKET ADJ / COUPON data rows


def _num(s):
    s = (s or "").replace(",", "").strip()
    return round(float(s), 2) if s else 0.0


# REMITTANCE columns as Metro sends them; the team's workbooks move Net in front of Gross,
# so columns are located by their header name and these positions are only the fallback.
DEFAULT_COLUMNS = {"control": 1, "store": 2, "date": 3, "invoice": 4, "gross": 5, "discount": 6, "net": 7}
HEADER_NAMES = {"control_number": "control", "control number": "control", "store number": "store",
                "invoice date": "date", "invoice": "invoice", "gross": "gross", "discount": "discount", "net": "net"}


def _columns(header_row):
    found = {}
    for i, h in enumerate(header_row):
        key = HEADER_NAMES.get(h.strip().lower())
        if key and key not in found:
            found[key] = i
    return {**DEFAULT_COLUMNS, **found} if {"gross", "discount", "net", "invoice"} <= set(found) else DEFAULT_COLUMNS


def _region_code(name):
    """'..._CA6-RMT_...csv' (Metro's own file name) or '... CA6.csv' (prepared test files)."""
    m = re.search(r"(?:^|[-_ ])(CA[56])(?=[-_ .]|$)", name.upper())
    return m.group(1) if m else None


def _file_date(name):
    """Metro's '_202608271334.csv' timestamp, or 'MM-DD-YYYY' in a prepared test file name."""
    if d := re.search(r"_(\d{8})\d{4}\.csv$", name, re.I):
        return datetime.strptime(d.group(1), "%Y%m%d").date()
    if d := re.search(r"(\d{2})-(\d{2})-(\d{4})", name):
        return datetime(int(d.group(3)), int(d.group(1)), int(d.group(2))).date()
    return None


def parse_metro(path):
    path = Path(path)
    region_code = _region_code(path.name)
    file_date = _file_date(path.name)

    headers, summaries, lines, other = {}, [], [], []
    cols = DEFAULT_COLUMNS
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row_no, row in enumerate(csv.reader(fh), start=1):
            if not row or not row[0].strip():
                continue
            rtype = row[0].strip().upper()
            if rtype not in headers:          # first row of a record type = header definition
                headers[rtype] = row
                if rtype == "REMITTANCE":
                    cols = _columns(row)
                continue
            if rtype == "REMITTANCE SUMMARY":
                summaries.append({"control": row[1].strip(), "vendor": row[2].strip() if len(row) > 2 else "",
                                  "name": row[3].strip() if len(row) > 3 else "",
                                  "total": _num(row[4] if len(row) > 4 else "0"), "row": row_no})
            elif rtype == "REMITTANCE":
                row = row + [""] * (8 - len(row))
                date_txt = row[cols["date"]].strip()
                lines.append({
                    "Row": row_no, "Control": row[cols["control"]].strip(), "Store": row[cols["store"]].strip(),
                    "Invoice Date": datetime.strptime(date_txt, "%Y%m%d").date() if date_txt else None,
                    "Invoice": row[cols["invoice"]].strip(), "Gross": _num(row[cols["gross"]]),
                    "Discount": _num(row[cols["discount"]]), "Net": _num(row[cols["net"]]),
                })
            else:
                rec = {"Row": row_no, "Record Type": rtype}
                hdr = headers[rtype]
                for i, v in enumerate(row[1:], start=1):
                    rec[hdr[i].strip() if i < len(hdr) and hdr[i].strip() else f"Col{i + 1}"] = v
                other.append(rec)
    return MetroRemittance(path=path, region_code=region_code, file_date=file_date, summaries=summaries,
                           lines=pd.DataFrame(lines), other=pd.DataFrame(other))
