"""Costco payment advice PDF -> output workbook.

Example:
    python automation/run_costco.py --remit "Costco Payment.PDF" ^
        --open-items "Open Item Reports/open item report costco.xlsx"
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import CODES_DIR  # noqa: E402
from costco.engine import run_costco  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="Costco cash application automation")
    ap.add_argument("--remit", required=True, help="Costco payment advice PDF")
    ap.add_argument("--open-items", required=True, help="Oracle Open Item Report, all regions (.xlsx)")
    ap.add_argument("--work-date", help="YYYY-MM-DD transaction date (default: advice date)")
    ap.add_argument("--rules", help="Rules workbook (default: Set of rules for automation/...)")
    ap.add_argument("--out", default=str(CODES_DIR / "Output"), help="Output folder")
    ap.add_argument("--force", action="store_true", help="Write output even if the PDF total does not tie")
    a = ap.parse_args()
    wd = datetime.strptime(a.work_date, "%Y-%m-%d").date() if a.work_date else None
    r = run_costco(a.remit, a.open_items, a.out, a.rules, wd, a.force)
    c = r["controls"]
    print(f"Payment {r['payment']}: total {r['total']:,.2f}")
    print("  Status counts: " + ", ".join(f"{k} {v}" for k, v in sorted(r["status_counts"].items())))
    for region, (ref, amt) in r["nac"].items():
        print(f"  {region}: {ref} = {amt:,.2f}")
    print("  Controls: " + ("ALL PASS" if c.all_pass else "FAILED -> " + "; ".join(c.failed())))
    print(f"  Output: {r['path']}")


if __name__ == "__main__":
    main()
