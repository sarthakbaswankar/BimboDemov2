"""Metro cash application: match, classify, build Upload / Receipt / control sheets."""
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from common.controls import Controls
from common.excel_writer import CONTROL_SHEET_WIDTHS
from common.excel_writer import write_workbook
from common.keys import clean, metro_key, oracle_invoice_text, significant
from common.open_items import load_open_items
from common.references import RefRegistry
from common.upload import receipt_frame, upload_frame
from config import load_params, rules_label
from metro.parser import parse_metro
from metro.summary import build_metro_summary


def _within(amount, tol, inclusive):
    return abs(amount) <= tol + 1e-9 if inclusive else abs(amount) < tol - 1e-9


def load_closed_trx(path):
    """Optional export of closed / all transactions -> set of normalised keys."""
    if not path:
        return None
    path = Path(path)
    df = pd.read_csv(path, dtype=object) if path.suffix.lower() == ".csv" else pd.read_excel(path, dtype=object)
    col = "Trx Number" if "Trx Number" in df.columns else df.columns[0]
    return {metro_key(v) for v in df[col].dropna()}


def process_control(rem, summary, lines, oi_all, oi_group, P, region, receipt_date, closed_keys, meta):
    pm, pc = P["METRO"], P["COMMON"]
    eps, tol, incl = pc["AMOUNT_EPSILON"], pm["WO_TOLERANCE"], pm["WO_INCLUSIVE"]
    control = summary["control"]
    ctl7 = control[-pm["CONTROL_DIGITS"]:]
    bill_to = pm["BILLTO_" + region.upper()]
    registry = RefRegistry(oi_all.refs, pc["SUFFIX_STYLE"], receipt_date)
    used = set()
    out = []

    for ln in lines.to_dict("records"):
        inv = clean(ln["Invoice"])
        key = metro_key(inv)
        rec = {**ln, "Match Key": key, "Oracle Trx": None, "Oracle Date": None, "Oracle Amount": None,
               "Oracle Bill-to": None, "Oracle Party": None, "Variance": None, "Status": "", "Category": "",
               "Reference": "", "Upload Amount": None, "WO Amount": None, "Receipt Amount": None,
               "Rule": "", "Note": ""}
        cands = [r for r in oi_group.by_key.get(key, []) if r["id"] not in used]
        if cands:
            doc = next((r for r in cands if abs(r["due"] - ln["Gross"]) < eps), cands[0])
            used.add(doc["id"])
            var = round(ln["Gross"] - doc["due"], 2)
            rec.update({"Oracle Trx": oracle_invoice_text(doc["trx"]), "Oracle Date": doc["date"],
                        "Oracle Amount": doc["due"], "Oracle Bill-to": doc["bill_to"], "Oracle Party": doc["party"],
                        "Variance": var, "Category": "Invoice"})
            if abs(var) < eps:
                rec.update(Status="Matched", Rule="M-03 Key match, gross = Oracle open amount",
                           **{"Receipt Amount": doc["due"]})
            elif _within(var, tol, incl):
                rec.update(Status="WO", Rule=f"M-08 Matched with variance within +/-{tol:g}",
                           **{"Receipt Amount": doc["due"], "WO Amount": var})
            else:
                rec.update(Status="Review", Rule="M-03 Matched with variance above tolerance",
                           Note="Partial payment applied at the paid gross amount; investigate the variance",
                           **{"Receipt Amount": ln["Gross"]})
            out.append(rec)
            continue

        # ---- not found (N/A) ------------------------------------------------
        if oi_group.by_key.get(key):
            rec["Note"] = "Invoice already applied earlier in this remittance (duplicate). "
        else:
            other = oi_all.by_key.get(key)
            if other:
                rec["Note"] = (f"Number exists on another customer ({other[0]['party']} {other[0]['bill_to']}); "
                               "not matched. ")
        sig = significant(inv) if inv.isdigit() else inv
        if inv.isdigit() and len(sig) < pm["INVOICE_MIN_DIGITS"]:
            rec["Category"] = "Deduction"
            if _within(ln["Net"], tol, incl):
                rec.update(Status="WO", Rule=f"M-07 Deduction within +/-{tol:g} -> write-off",
                           **{"WO Amount": ln["Net"]})
            else:
                ref, suffixed = registry.make(sig)
                rec.update(Status="Deduction", Reference=ref, Rule="M-05 Fewer than 10 digits and not in Oracle",
                           **{"Upload Amount": round(-ln["Net"], 2)})
                if suffixed:
                    rec["Note"] += "Deduction number already exists in Oracle; suffix added. "
            out.append(rec)
            continue

        # invoice-shaped -> SB or NF (never written off)
        basis = ln["Gross"] if pm["DISC_BASIS"].upper() == "TOTAL" else ln["Net"]
        rec["Category"] = "Invoice not open"
        sb_base, nf_base = pm["SB_PREFIX"] + sig, pm["NF_PREFIX"] + sig
        if closed_keys is not None and key in closed_keys:
            ref, _ = registry.make(sb_base)
            rec.update(Status="SB", Rule="M-10 Exists in Oracle (closed-transactions export)")
        elif registry.exists(sb_base):
            ref, _ = registry.make(sb_base)
            rec.update(Status="SB", Rule="M-10 SB INV already exists in Oracle -> suffix")
        elif closed_keys is not None:
            ref, _ = registry.make(nf_base)
            rec.update(Status="NF", Rule="M-09 Not in Oracle (checked against closed-transactions export)")
        else:
            ref, _ = registry.make(nf_base)
            rec.update(Status="NF", Rule="M-09 Not open in Oracle; provisional NF")
            rec["Note"] += "Provisional NF: confirm in Oracle; change to SB INV- if the invoice exists (closed). "
        rec.update(Reference=ref, **{"Upload Amount": round(-basis, 2)})
        if not inv.isdigit():
            rec["Note"] += "Non-numeric invoice reference. "
        out.append(rec)

    fmt = pd.DataFrame(out)

    # ---- upload --------------------------------------------------------------
    up_lines = []
    for r in out:
        if r["Upload Amount"] is not None and r["Status"] in ("Deduction", "NF", "SB"):
            up_lines.append({"ref": r["Reference"], "amount": r["Upload Amount"], "category": r["Status"],
                             "source": r["Invoice"]})
    if pm["DISC_BASIS"].upper() == "TOTAL":
        disc = round(lines["Discount"].sum(), 2)
        disc_note = "Total remittance discount (DISC_BASIS=TOTAL)"
    else:
        disc = round(sum(r["Discount"] for r in out if r["Category"] == "Invoice"), 2)
        disc_note = "Discount on matched invoices (DISC_BASIS=MATCHED)"
    disc_ref = None
    if abs(disc) >= eps:
        disc_ref, _ = registry.make(pm["DISC_PREFIX"] + ctl7)
        up_lines.append({"ref": disc_ref, "amount": disc, "category": "DISC", "source": disc_note})
    nac_ref, _ = registry.make(pm["NAC_PREFIX"] + ctl7)
    nac_amount = round(-sum(x["amount"] for x in up_lines), 2)
    up_lines.append({"ref": nac_ref, "amount": nac_amount, "category": "NAC", "source": "Balancing credit/debit"})
    for x in up_lines:
        x.update(bill_to=bill_to, trx_date=receipt_date, region=region, group=nac_ref)
    upload = upload_frame(up_lines, pc, pm["TERMS_DR"], pm["TERMS_CR"])

    # ---- receipt -------------------------------------------------------------
    refs = [{"value": r["Oracle Trx"], "amount": r["Receipt Amount"], "region": region,
             "bill_to": r["Oracle Bill-to"], "category": r["Status"],
             "note": "" if r["Oracle Bill-to"] == bill_to else "Document on another Metro-group bill-to"}
            for r in out if r["Receipt Amount"] is not None]
    refs.append({"value": nac_ref, "amount": nac_amount, "region": region, "bill_to": bill_to,
                 "category": "NAC", "note": "Created by the transaction upload"})
    yymmdd = receipt_date.strftime("%y%m%d")
    header = {"number": nac_ref, "amount": summary["total"], "date": receipt_date, "method": pm["RECEIPT_METHOD"],
              "account": bill_to, "batch": pm["BATCH_NAME_PATTERN"].format(region=region, yymmdd=yymmdd)}
    receipt = receipt_frame(header, refs)

    # ---- controls -------------------------------------------------------------
    c = Controls()
    eps2 = 0.01
    c.amount("Remittance lines net = summary total", summary["total"], lines["Net"].sum(), eps2)
    c.amount("Gross - discount = net", lines["Net"].sum(), lines["Gross"].sum() - lines["Discount"].sum(), eps2)
    c.check("Every line has exactly one status", fmt["Status"].ne("").all() and len(fmt) == len(lines),
            f"{len(fmt)} classified of {len(lines)}")
    c.amount("Upload sums to 0", 0.0, upload["Transaction Line Amount"].sum(), eps2)
    wo_total = round(fmt["WO Amount"].fillna(0).sum(), 2)
    c.amount("Receipt references + write-offs = remittance total", summary["total"],
             receipt["Reference Amount"].sum() + wo_total, eps2,
             f"Receipt refs {receipt['Reference Amount'].sum():,.2f}; write-offs {wo_total:,.2f}")
    c.check("No open item applied twice", len(used) == fmt["Oracle Trx"].notna().sum())
    c.check("No duplicate upload reference", upload["Transaction Number"].is_unique)
    wo = fmt[fmt["Status"] == "WO"]
    c.check(f"Write-offs within +/-{tol:g}", all(_within(v, tol, incl) for v in wo["WO Amount"]),
            f"{len(wo)} write-off line(s)")

    review = fmt[(fmt["Status"] == "Review") | fmt["Note"].ne("")][
        ["Row", "Store", "Invoice Date", "Invoice", "Gross", "Net", "Status", "Reference", "Rule", "Note"]]

    summary_blocks = build_metro_summary({
        "fmt": fmt, "lines": lines, "upload": upload, "receipt": receipt, "controls": c, "review": review,
        "header": header, "meta": meta, "total": summary["total"], "region": region,
        "region_code": rem.region_code or "-", "control": control, "file": rem.path.name,
        "vendor": f"{summary['vendor']} {summary['name']}".strip(), "receipt_date": receipt_date,
        "nac_ref": nac_ref, "nac_amount": nac_amount, "disc_basis": pm["DISC_BASIS"],
    })

    formatted_cols = ["Row", "Store", "Invoice Date", "Invoice", "Net", "Gross", "Discount", "Match Key",
                      "Oracle Trx", "Oracle Date", "Oracle Amount", "Oracle Bill-to", "Oracle Party", "Variance",
                      "Status", "Category", "Reference", "Upload Amount", "Receipt Amount", "WO Amount", "Rule",
                      "Note"]
    fmt = fmt[formatted_cols]
    sheets = [
        {"name": "Summary", "blocks": summary_blocks},
        {"name": "Raw_Remittance", "df": lines},
        {"name": "Formatted", "df": fmt, "status_col": "Status"},
        {"name": "Upload", "df": upload, "status_col": "Category"},
        {"name": "Receipt", "df": receipt},
        {"name": "WriteOffs", "df": wo},
        {"name": "Deductions", "df": fmt[fmt["Category"] == "Deduction"].sort_values("Net"), "status_col": "Status"},
        {"name": "NF_SB", "df": fmt[fmt["Status"].isin(["NF", "SB"])], "status_col": "Status"},
        {"name": "Review", "df": review, "status_col": "Status"},
        {"name": "Control_Checks", "widths": CONTROL_SHEET_WIDTHS,
         "blocks": c.sheet_blocks(f"Control Checks: Metro {region}, control {control}",
                                  f"Remittance {rem.path.name}  |  receipt date {receipt_date:%Y-%m-%d}")},
    ]
    # sheet order: Summary, Raw_Remittance, Formatted, Upload, Receipt, WriteOffs, then the other sheets
    if rem.other is not None and len(rem.other):
        sheets.insert(8, {"name": "Other_Records", "df": rem.other})
    stats = {"control": control, "total": summary["total"], "nac": nac_amount, "nac_ref": nac_ref,
             "disc": disc, "status_counts": fmt["Status"].value_counts().to_dict(), "controls": c,
             "formatted": fmt, "upload": upload, "receipt": receipt}
    return sheets, stats


def run_metro(remit_path, oi_path, out_dir, rules_path=None, closed_trx=None, receipt_date=None, region=None):
    P = load_params(rules_path)
    pm = P["METRO"]
    rem = parse_metro(remit_path)
    if region is None:
        if not rem.region_code or f"REGION_{rem.region_code}" not in pm:
            raise ValueError("Region not found in the file name (CA5/CA6); pass --region Ontario|Quebec")
        region = pm[f"REGION_{rem.region_code}"]
    receipt_date = receipt_date or rem.file_date or date.today()
    oi_all = load_open_items(oi_path, metro_key)
    oi_group = oi_all.subset(pm["GROUP_BILLTOS"])
    closed_keys = load_closed_trx(closed_trx)
    meta = {"oi_file": Path(oi_path).name, "oi_rows": len(oi_all), "group_rows": len(oi_group),
            "closed": Path(closed_trx).name if closed_trx else "", "rules": rules_label(rules_path),
            "run_at": datetime.now()}
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    for summary in rem.summaries:
        lines = rem.lines[rem.lines["Control"] == summary["control"]].reset_index(drop=True)
        sheets, stats = process_control(rem, summary, lines, oi_all, oi_group, P, region, receipt_date,
                                        closed_keys, meta)
        path = out_dir / f"Metro_{region}_{summary['control']}_{receipt_date:%Y%m%d}.xlsx"
        write_workbook(path, sheets)
        stats["path"] = path
        results.append(stats)
    return results
