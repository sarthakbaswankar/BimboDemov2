"""Costco payment advice PDF parser (replaces Ctrl+A / Ctrl+C / Text-to-Columns)."""
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber

AMT = r"-?[\d,]+\.\d{2}"
LINE_RE = re.compile(rf"^(?P<inv>\S+)\s+(?P<po>.*?)\s*(?P<date>\d{{2}}/\d{{2}}/\d{{4}})\s+"
                     rf"(?P<gross>{AMT})\s+(?P<disc>{AMT})\s+(?P<amt>{AMT})$")
DATE_RE = re.compile(r"\d{2}/\d{2}/\d{4}")


@dataclass
class CostcoAdvice:
    path: Path
    advice_date: object = None
    payment: str = ""
    currency: str = ""
    vendor: str = ""
    total: float | None = None
    pages: int = 0
    lines: pd.DataFrame = None
    exceptions: list = field(default_factory=list)


def _amt(s):
    return round(float(s.replace(",", "")), 2)


def parse_costco_pdf(path):
    path = Path(path)
    adv = CostcoAdvice(path=path)
    rows = []
    with pdfplumber.open(path) as pdf:
        adv.pages = len(pdf.pages)
        for pno, page in enumerate(pdf.pages, start=1):
            for raw in (page.extract_text() or "").splitlines():
                line = raw.strip()
                if not line:
                    continue
                if m := re.match(r"^Date\s*:\s*(\d{2}/\d{2}/\d{4})", line):
                    adv.advice_date = datetime.strptime(m.group(1), "%m/%d/%Y").date()
                    continue
                if m := re.match(r"^Payment\s*#\s*:\s*(\d+)\s*(\w+)?", line):
                    adv.payment, adv.currency = m.group(1), m.group(2) or ""
                    continue
                if m := re.match(r"^Vendor#\s*-\s*(\S+)", line):
                    adv.vendor = m.group(1)
                    continue
                if m := re.match(rf"^Remittance TOTAL\s+({AMT})", line, re.I):
                    adv.total = _amt(m.group(1))
                    continue
                if m := LINE_RE.match(line):
                    rows.append({
                        "Page": pno, "Invoice": m["inv"], "PO / Description": m["po"].strip(),
                        "Date": datetime.strptime(m["date"], "%m/%d/%Y").date(),
                        "Gross": _amt(m["gross"]), "Discount": _amt(m["disc"]), "Amount": _amt(m["amt"]),
                    })
                elif DATE_RE.search(line):
                    adv.exceptions.append({"Page": pno, "Text": line})
    adv.lines = pd.DataFrame(rows)
    return adv
