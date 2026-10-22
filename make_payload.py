"""Builds the JSON payload for POST /process_national_customers from a remittance file.

Usage (from this folder):
    python make_payload.py costco "Costco Payment.PDF"
    python make_payload.py metro "CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv" --output json

Writes payloads/<file name>.json, ready to paste into Postman (Body -> raw -> JSON) or send with curl:
    curl -X POST http://127.0.0.1:8000/process_national_customers -H "Content-Type: application/json" ^
         --data-binary "@payloads/Costco Payment.json" -o response.json
"""
import argparse
import base64
import json
from pathlib import Path


def build(customer, path, output="excel"):
    path = Path(path)
    return {"customer": customer, "filename": path.name,
            "file_base64": base64.b64encode(path.read_bytes()).decode("ascii"), "output": output}


def main():
    ap = argparse.ArgumentParser(description="Build the JSON payload for POST /process_national_customers")
    ap.add_argument("customer", help="metro or costco")
    ap.add_argument("file", help="Metro .csv or Costco .pdf")
    ap.add_argument("--output", default="base64", choices=["base64", "json", "excel"], help="what the API should return")
    a = ap.parse_args()
    out = Path(__file__).resolve().parent / "payloads" / f"{Path(a.file).stem}.json"
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps(build(a.customer, a.file, a.output), indent=2), encoding="utf-8")
    print(f"Payload written: {out}  ({out.stat().st_size / 1024:,.0f} KB)")


if __name__ == "__main__":
    main()
