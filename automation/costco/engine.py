"""Costco cash application: combine, match, resolve, split by region, build outputs."""
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

import pandas as pd

from common.controls import Controls
from common.excel_writer import CONTROL_SHEET_WIDTHS
from common.excel_writer import write_workbook
from common.keys import clean, oracle_invoice_text, r12, significant
from common.open_items import load_open_items
from common.references import RefRegistry
from common.upload import receipt_frame, upload_frame
from config import REGION_LETTERS, load_params, rules_label
from costco.summary import build_costco_summary
from costco.pdf_parser import parse_costco_pdf
from costco.resolver import Resolver

REGIONS = ["Atlantic", "Western", "Ontario", "Quebec"]
STATUS_ORDER = {"Deduction": 0, "Matched": 1, "Resolved": 2, "WO": 3, "SP": 4, "OP": 5, "NF": 6, "Review": 7}


def store_from_po(po):
    """'006560803000' -> '656' (PO = store + MMDD + 000)."""
    po = po.replace(" ", "")
    if len(po) == 12 and po.isdigit():
        return significant(po[:-7])
    return ""


def digits_region(number):
    """Team fallback rule on the number as printed with leading zeros stripped."""
    sig = significant(number) if number.isdigit() else number
    if number.isdigit() and len(sig) == 12:
        return {"1": "Atlantic", "2": "Western", "3": "Quebec"}.get(sig[0], "Ontario")
    return "Ontario"


def is_deduction(inv, pk):
    s = inv.upper()
    if s.startswith("RV"):
        return True
    if not s.isdigit():
        return False
    sig = significant(s)
    if len(s) >= pk["DEDUCTION_MIN_DIGITS"]:
        return True
    return any(sig.startswith(p) for p in pk["DEDUCTION_PREFIXES"] if p.isdigit())


def run_costco(pdf_path, oi_path, out_dir, rules_path=None, work_date=None, force=False):
    P = load_params(rules_path)
    pc, pk = P["COMMON"], P["COSTCO"]
    eps = pc["AMOUNT_EPSILON"]
    adv = parse_costco_pdf(pdf_path)
    raw = adv.lines
    pdf_sum = round(raw["Amount"].sum(), 2)
    if adv.total is None or abs(pdf_sum - adv.total) > 0.01:
        msg = f"Parsed lines total {pdf_sum:,.2f} does not equal Remittance TOTAL {adv.total}"
        if not force:
            raise ValueError(msg + " (hard stop; use --force to write the workbook anyway)")
    work_date = work_date or adv.advice_date

    billto_region = {pk["BILLTO_ATLANTIC"]: "Atlantic", pk["BILLTO_WESTERN"]: "Western",
                     pk["BILLTO_ONTARIO"]: "Ontario", pk["BILLTO_QUEBEC"]: "Quebec"}
    for b in pk["BILLTO_ONTARIO_EXTRA"]:
        billto_region[b] = "Ontario"
    region_billto = {r: pk["BILLTO_" + r.upper()] for r in REGIONS}

    oi = load_open_items(oi_path, r12)
    registry = RefRegistry(oi.refs, pc["SUFFIX_STYLE"], work_date)

    # ---- C2 combine duplicate invoice numbers --------------------------------
    groups = OrderedDict()
    for rec in raw.to_dict("records"):
        inv = clean(rec["Invoice"]).upper()
        key = r12(inv)
        g = groups.setdefault(key, {"key": key, "invoice": inv, "po": rec["PO / Description"], "date": rec["Date"],
                                    "paid": 0.0, "parts": [], "pages": set(), "dates": set(), "notes": []})
        g["paid"] = round(g["paid"] + rec["Amount"], 2)
        g["parts"].append(rec["Amount"])
        g["pages"].add(rec["Page"])
        g["dates"].add(rec["Date"])
    lines = list(groups.values())
    for ln in lines:
        ln["store"] = store_from_po(ln["po"])
        ln["audit"] = " + ".join(f"{a:,.2f}" for a in ln["parts"]) if len(ln["parts"]) > 1 else ""
        if len(ln["dates"]) > 1:
            ln["notes"].append("Combined lines have different dates")

    # ---- C3 match: reserve every primary document first ---------------------
    used = set()
    for ln in lines:
        cands = [d for d in oi.by_key.get(ln["key"], []) if d["id"] not in used]
        ln["primary"] = min(cands, key=lambda d: abs(d["due"] - ln["paid"])) if cands else None
        if ln["primary"]:
            used.add(ln["primary"]["id"])

    # ---- C4/C5b classify and resolve ------------------------------------------
    resolver = Resolver(oi, used, P)
    order = sorted(range(len(lines)), key=lambda i: (lines[i]["date"],
                                                     lines[i]["paid"] - (lines[i]["primary"] or {"due": 0})["due"],
                                                     lines[i]["key"]))
    log_rows = []
    for i in order:
        ln = lines[i]
        p = ln["primary"]
        ln.update(added=[], var=None, L=None, status="", rule="", ref="", upload=None, wo=None, region="")
        if p:
            ln["var0"] = round(ln["paid"] - p["due"], 2)
            if abs(ln["var0"]) < eps:
                ln.update(status="Matched", rule="C-03 R12 match, paid = open", var=0.0, L=0.0)
                continue
            added, log, var, notes = resolver.resolve(ln["key"], ln["paid"], [p], {p["ship_to"]} - {""},
                                                      p["date"] or ln["date"])
        elif is_deduction(ln["invoice"], pk):
            ln.update(status="Deduction", region="Ontario", var0=ln["paid"])
            continue
        else:
            ln["var0"] = ln["paid"]
            ship_tos = oi.store_shiptos.get(ln["store"], set()) if ln["store"] else set()
            added, log, var, notes = (resolver.resolve(ln["key"], ln["paid"], [], ship_tos, ln["date"])
                                      if ship_tos else ([], [], ln["paid"], []))
        ln["added"], ln["var"], ln["L"] = added, var, round(-var, 2)
        ln["notes"] += notes
        for e in log:
            log_rows.append({"Invoice": ln["invoice"], "Store": ln["store"], "Paid": ln["paid"], **e})
        if not p and not added:
            ln["status"] = "NF"
            continue
        if abs(var) < eps:
            ln.update(status="Resolved", rule=log[-1]["Rule"] if log else "")
        elif resolver.within_tol(var):
            ln.update(status="WO", wo=var, rule=f"C-15 Residual within +/-{pk['WO_TOLERANCE']:g} -> write-off")
        else:
            ln.update(status="SP" if ln["L"] > 0 else "OP",
                      rule=f"C-16 Residual above +/-{pk['WO_TOLERANCE']:g} -> " + ("short pay" if ln["L"] > 0
                                                                                   else "overpay"))
        if log and ln["status"] != "Resolved":
            ln["rule"] = "; ".join(sorted({e["Rule"].split(":")[0] for e in log})) + " | " + ln["rule"]
        if not p and added:
            ln["notes"].append("Invoice number not in Oracle; documents found by the engine, please verify")
        if ln["status"] in ("SP", "OP") and abs(ln["L"]) > pk["REVIEW_RESIDUAL_ABOVE"]:
            ln["notes"].append(f"Large {ln['status']} ({ln['L']:,.2f}); check for a missing invoice/credit "
                               "before uploading")

    # ---- C5 regions, C7 references and upload amounts -------------------------
    for ln in lines:
        docs = ([ln["primary"]] if ln["primary"] else []) + ln["added"]
        if not ln["region"]:
            bt = next((d["bill_to"] for d in docs if d["bill_to"] in billto_region), None)
            if bt:
                ln["region"] = billto_region[bt]
            else:
                ln["region"] = digits_region(ln["invoice"])
                ln["rule"] = (ln["rule"] + " | " if ln["rule"] else "") + "C-07 Region by number pattern"
                store_regions = {billto_region.get(b) for b in oi.store_billtos.get(ln["store"], set())} - {None}
                if len(store_regions) == 1 and ln["region"] not in store_regions:
                    ln["notes"].append(f"Store {ln['store']} belongs to {store_regions.pop()} in Oracle; "
                                       f"number pattern says {ln['region']}")
        sig = significant(ln["invoice"]) if ln["invoice"].isdigit() else ln["invoice"]
        nf_num = r12(ln["invoice"]) if (pk["NF_LEADING_ZERO"] or len(sig) > 12) else sig
        if ln["status"] == "Deduction":
            if ln["invoice"].startswith("RV"):
                ln["ref"], _ = registry.make(ln["invoice"])
                ln["rule"] = "C-05 RV deduction -> Ontario, uploaded under its RV number"
            else:
                ln["ref"], _ = registry.make("NF-" + (significant(nf_num) if not pk["NF_LEADING_ZERO"] else nf_num))
                ln["rule"] = "C-06 Deduction pattern (65/14/604/13+ digits) not in Oracle -> NF, Ontario"
            ln["upload"] = round(-ln["paid"], 2)
        elif ln["status"] == "NF":
            ln["ref"], _ = registry.make("NF-" + nf_num)
            ln["rule"] = (ln["rule"] + " | " if ln["rule"] else "") + "C-09 Invoice not in open items -> NF"
            ln["upload"] = round(-ln["paid"], 2)
            ln["L"] = round(-ln["paid"], 2)
        elif ln["status"] in ("SP", "OP"):
            base = f"{ln['store'] or ln['key']}-{ln['date']:%m%d%y}-{ln['status']}"
            ln["ref"], suffixed = registry.make(base)
            if suffixed:
                ln["notes"].append(f"{base} already used; suffix added")
            ln["upload"] = ln["L"]

    # ---- upload: per-region lines + NAC balancing line ------------------------
    up_lines, nac_amounts = [], {}
    for region in REGIONS:
        rl = [ln for ln in lines if ln["region"] == region and ln["upload"] is not None]
        if not rl:
            continue
        nac_ref, _ = registry.make(pk["NAC_PATTERN"].format(payment=adv.payment, letter=REGION_LETTERS[region]))
        for ln in rl:
            up_lines.append({"ref": ln["ref"], "amount": ln["upload"], "category": ln["status"], "region": region,
                             "group": nac_ref, "source": ln["invoice"]})
        nac_amt = round(-sum(ln["upload"] for ln in rl), 2)
        nac_amounts[region] = (nac_ref, nac_amt)
        up_lines.append({"ref": nac_ref, "amount": nac_amt, "category": "NAC", "region": region,
                         "group": nac_ref, "source": "Regional balancing line"})
    for x in up_lines:
        x.update(bill_to=region_billto[x["region"]], trx_date=work_date)
    upload = upload_frame(up_lines, pc, pk["TERMS_DR"], pk["TERMS_CR"])

    # ---- receipt ----------------------------------------------------------------
    refs = []
    for ln in lines:
        for d in ([ln["primary"]] if ln["primary"] else []) + ln["added"]:
            refs.append({"value": oracle_invoice_text(d["trx"]), "amount": d["due"], "region": ln["region"],
                         "bill_to": d["bill_to"], "category": ln["status"],
                         "note": "" if d is ln["primary"] else "Attached by resolution engine"})
    for region, (nac_ref, nac_amt) in nac_amounts.items():
        refs.append({"value": nac_ref, "amount": nac_amt, "region": region, "bill_to": region_billto[region],
                     "category": "NAC", "note": "Created by the transaction upload"})
    header = {"number": adv.payment, "amount": adv.total or pdf_sum, "date": work_date,
              "method": pk["RECEIPT_METHOD"], "account": pk["BILLTO_ONTARIO"],
              "batch": f"Costco {work_date:%y%m%d}"}
    receipt = receipt_frame(header, refs)

    # ---- sheets -------------------------------------------------------------------
    def line_rows(ln):
        p = ln["primary"]
        k = round(ln["paid"] - p["due"], 2) if p else ln["paid"]
        rows = [{
            "Invoice": ln["invoice"], "PO": ln["po"], "Date": ln["date"], "Paid": ln["paid"], "Store": ln["store"],
            "R12": ln["key"], "Bill-to": p["bill_to"] if p else None, "Ship-to": p["ship_to"] if p else None,
            "Oracle Trx": oracle_invoice_text(p["trx"]) if p else None, "Oracle Date": p["date"] if p else None,
            "Open Amount": p["due"] if p else None, "Var (K)": k, "Net Var (L)": ln["L"],
            "Status": ln["status"], "Reference": ln["ref"], "Upload Amount": ln["upload"],
            "Region": ln["region"], "Rule": ln["rule"], "Combined Lines": ln["audit"], "Note": "; ".join(ln["notes"]),
        }]
        for d in ln["added"]:
            rows.append({"Invoice": oracle_invoice_text(d["trx"]), "Paid": None, "Store": ln["store"],
                         "R12": d["key"], "Bill-to": d["bill_to"], "Ship-to": d["ship_to"],
                         "Oracle Trx": oracle_invoice_text(d["trx"]), "Oracle Date": d["date"],
                         "Open Amount": d["due"], "Var (K)": round(-d["due"], 2), "Status": "Attached",
                         "Region": ln["region"], "Rule": f"attached to {ln['invoice']}"})
        return rows

    def sort_key(ln):
        return (STATUS_ORDER.get(ln["status"], 9), ln["var"] if ln["var"] is not None else ln["paid"], ln["key"])

    region_sheets = []
    for region in REGIONS:
        rows = []
        for ln in sorted((x for x in lines if x["region"] == region), key=sort_key):
            rows += line_rows(ln)
        region_sheets.append({"name": region, "df": pd.DataFrame(rows), "status_col": "Status"})

    all_rows = pd.DataFrame([line_rows(ln)[0] for ln in lines])
    combined = all_rows[["Invoice", "PO", "Date", "Paid", "Store", "R12", "Combined Lines", "Region", "Status"]]
    wo_df = all_rows[all_rows["Status"] == "WO"].assign(**{"WO Amount": [ln["wo"] for ln in lines
                                                                         if ln["status"] == "WO"]})
    review = all_rows[all_rows["Note"].ne("")]
    if adv.exceptions:
        review = pd.concat([review, pd.DataFrame([{"Invoice": "(parse exception)", "Note": f"Page {e['Page']}: "
                                                   f"{e['Text']}", "Status": "Review"} for e in adv.exceptions])])

    # ---- controls ----------------------------------------------------------------
    c = Controls()
    c.amount("PDF lines = Remittance TOTAL", adv.total or 0.0, pdf_sum, 0.01, f"{len(raw)} lines on {adv.pages} pages")
    c.amount("Combined lines total = Remittance TOTAL", adv.total or 0.0, sum(ln["paid"] for ln in lines), 0.01,
             f"{len(raw)} lines -> {len(lines)} after combining duplicates")
    c.check("Every line has exactly one status", all(ln["status"] for ln in lines))
    for region, (nac_ref, _) in nac_amounts.items():
        c.amount(f"Upload {nac_ref} sums to 0", 0.0,
                 upload.loc[upload["NAC Group"] == nac_ref, "Transaction Line Amount"].sum(), 0.01)
    wo_total = round(sum(ln["wo"] for ln in lines if ln["status"] == "WO"), 2)
    c.amount("Receipt references + write-offs = remittance total", adv.total or 0.0,
             receipt["Reference Amount"].sum() + wo_total, 0.01,
             f"Receipt refs {receipt['Reference Amount'].sum():,.2f}; write-offs {wo_total:,.2f}")
    doc_ids = [d["id"] for ln in lines for d in ([ln["primary"]] if ln["primary"] else []) + ln["added"]]
    c.check("No open item used twice", len(doc_ids) == len(set(doc_ids)), f"{len(doc_ids)} documents applied")
    bad12 = [r["value"] for r in refs if r["category"] != "NAC" and r["value"].isdigit()
             and len(r["value"]) != 12 and not is_deduction(r["value"], pk)]
    c.check("Receipt invoice numbers are 12 digits", not bad12, ", ".join(bad12[:10]))
    c.check(f"Write-offs within +/-{pk['WO_TOLERANCE']:g}",
            all(resolver.within_tol(ln["wo"]) for ln in lines if ln["status"] == "WO"))
    c.check("No duplicate upload reference", upload["Transaction Number"].is_unique)

    # ---- summary -------------------------------------------------------------------
    meta = {"oi_file": Path(oi_path).name, "oi_rows": len(oi), "rules": rules_label(rules_path),
            "run_at": datetime.now()}
    summary_blocks = build_costco_summary({
        "adv": adv, "raw": raw, "lines": lines, "upload": upload, "receipt": receipt, "controls": c,
        "review": review, "meta": meta, "header": header, "region_billto": region_billto, "nac_amounts": nac_amounts,
        "log": pd.DataFrame(log_rows), "work_date": work_date, "wo_total": wo_total,
    })

    # sheet order: Summary, Raw_Remittance, Formatted, Upload, Receipt, WriteOffs, then the other sheets
    sheets = [{"name": "Summary", "blocks": summary_blocks},
              {"name": "Raw_Remittance", "df": raw},
              {"name": "Formatted", "df": combined, "status_col": "Status"},
              {"name": "Upload", "df": upload, "status_col": "Category"},
              {"name": "Receipt", "df": receipt},
              {"name": "WriteOffs", "df": wo_df}] + region_sheets + [
              {"name": "Resolution_Log", "df": pd.DataFrame(log_rows)},
              {"name": "Review", "df": review, "status_col": "Status"},
              {"name": "Control_Checks", "widths": CONTROL_SHEET_WIDTHS,
               "blocks": c.sheet_blocks(f"Control Checks: Costco payment {adv.payment}",
                                        f"Payment advice {adv.path.name}  |  work date {work_date:%Y-%m-%d}")}]
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"Costco_{adv.payment}_{work_date:%Y%m%d}.xlsx"
    write_workbook(path, sheets)
    return {"path": path, "payment": adv.payment, "total": adv.total, "controls": c, "lines": lines,
            "upload": upload, "receipt": receipt, "nac": nac_amounts, "log": pd.DataFrame(log_rows),
            "status_counts": pd.Series([ln["status"] for ln in lines]).value_counts().to_dict()}
