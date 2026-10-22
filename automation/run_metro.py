"""Metro remittance -> output workbook.

Example:
    python automation/run_metro.py --remit "CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv" ^
        --open-items "Open Item Reports/Open Item report Metro.xlsx"
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import CODES_DIR  # noqa: E402
from metro.engine import run_metro  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="Metro cash application automation")
    ap.add_argument("--remit", required=True, help="Metro remittance CSV (CA5 = Quebec, CA6 = Ontario)")
    ap.add_argument("--open-items", required=True, help="Oracle Open Item Report (.xlsx)")
    ap.add_argument("--closed-trx", help="Optional export of closed/all transactions to decide NF vs SB")
    ap.add_argument("--receipt-date", help="YYYY-MM-DD (default: date in the remittance file name)")
    ap.add_argument("--region", choices=["Ontario", "Quebec"], help="Override region from the file name")
    ap.add_argument("--rules", help="Rules workbook (default: Set of rules for automation/...)")
    ap.add_argument("--out", default=str(CODES_DIR / "Output"), help="Output folder")
    a = ap.parse_args()
    rd = datetime.strptime(a.receipt_date, "%Y-%m-%d").date() if a.receipt_date else None
    for r in run_metro(a.remit, a.open_items, a.out, a.rules, a.closed_trx, rd, a.region):
        c = r["controls"]
        print(f"Control {r['control']}: total {r['total']:,.2f} | NAC {r['nac_ref']} {r['nac']:,.2f}")
        print("  Status counts: " + ", ".join(f"{k} {v}" for k, v in sorted(r["status_counts"].items())))
        print("  Controls: " + ("ALL PASS" if c.all_pass else "FAILED -> " + "; ".join(c.failed())))
        print(f"  Output: {r['path']}")


if __name__ == "__main__":
    main()
