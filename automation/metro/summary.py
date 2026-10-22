"""Detailed Summary sheet for a Metro run (sectioned layout, see excel_writer._write_blocks).

Line-level details live in their own sheets (Formatted, Deductions, NF_SB, Upload,
Receipt); the control checks have their own Control_Checks sheet.
"""
import pandas as pd

MEANING = {
    "Matched": "Invoice open in Oracle, amount agrees: applied on the receipt",
    "WO": "Within +/-$10: not uploaded, written off on the receipt",
    "Deduction": "Metro deduction not in Oracle: uploaded at net x -1",
    "NF": "Invoice not found in Oracle: uploaded as NF- at x -1",
    "SB": "Invoice exists but is closed: uploaded as SB INV- at x -1",
    "Review": "Needs a decision before upload (see Review sheet)",
}
ORDER = ["Matched", "WO", "Deduction", "NF", "SB", "Review"]


def banner(controls, review_count):
    n = len(controls.rows)
    if not controls.all_pass:
        return (f"BLOCKED: {len(controls.failed())} of {n} control checks failed (see Control_Checks sheet). "
                "Do not upload until fixed.", "BLOCKED")
    if review_count:
        return (f"READY AFTER REVIEW: all {n} control checks PASS (see Control_Checks sheet). {review_count} line(s) "
                "on the Review sheet need a decision before upload.", "REVIEW")
    return f"READY TO UPLOAD: all {n} control checks PASS and nothing to review.", "READY"


def build_metro_summary(ctx):
    fmt, lines, receipt = ctx["fmt"], ctx["lines"], ctx["receipt"]
    c, review, hdr, meta = ctx["controls"], ctx["review"], ctx["header"], ctx["meta"]
    upload = ctx["upload"]
    total = ctx["total"]
    n = len(fmt)

    # ---- 3 classification -------------------------------------------------------
    rows = []
    for s in ORDER:
        g = fmt[fmt["Status"] == s]
        if not len(g):
            continue
        rows.append({"Status": s, "Meaning": MEANING[s], "Lines": len(g), "% of lines": len(g) / n,
                     "Gross": round(g["Gross"].sum(), 2), "Discount": round(g["Discount"].sum(), 2),
                     "Net (paid)": round(g["Net"].sum(), 2),
                     "Upload amount": round(g["Upload Amount"].fillna(0).sum(), 2),
                     "On receipt": round(g["Receipt Amount"].fillna(0).sum(), 2)})
    classes = pd.DataFrame(rows)

    # ---- 4 bridges ----------------------------------------------------------------
    matched_applied = round(fmt["Receipt Amount"].fillna(0).sum(), 2)
    wo_total = round(fmt["WO Amount"].fillna(0).sum(), 2)
    cash = round(matched_applied + ctx["nac_amount"] + wo_total, 2)
    receipt_bridge = pd.DataFrame([
        {"Step": "Matched invoices applied at Oracle amount", "Amount": matched_applied},
        {"Step": f"{ctx['nac_ref']} balancing credit (from the upload)", "Amount": ctx["nac_amount"]},
        {"Step": "Write-offs (not uploaded)", "Amount": wo_total},
        {"Step": "Cash received (should equal remittance total)", "Amount": cash},
        {"Step": "Remittance total", "Amount": total},
        {"Step": "Difference", "Amount": round(cash - total, 2)},
    ])
    up = upload[upload["Category"] != "NAC"]
    nac_rows = [{"Component": cat, "Lines": int((up["Category"] == cat).sum()),
                 "Upload amount": round(up.loc[up["Category"] == cat, "Transaction Line Amount"].sum(), 2)}
                for cat in ["Deduction", "NF", "SB", "DISC"] if (up["Category"] == cat).any()]
    nac_rows.append({"Component": "Sum of upload lines", "Lines": len(up),
                     "Upload amount": round(up["Transaction Line Amount"].sum(), 2)})
    nac_rows.append({"Component": f"{ctx['nac_ref']} = sum x -1", "Lines": 1, "Upload amount": ctx["nac_amount"]})
    nac_build = pd.DataFrame(nac_rows)

    # ---- 6 action items ---------------------------------------------------------------
    m = fmt[fmt["Oracle Trx"].notna()]
    actions = []
    prov = fmt["Note"].str.contains("Provisional NF", na=False).sum()
    if prov:
        actions.append(("Confirm provisional NF lines in Oracle: if the invoice exists (closed) change to SB INV-",
                        int(prov), "Review / NF_SB"))
    var_rev = (fmt["Status"] == "Review").sum()
    if var_rev:
        actions.append(("Investigate matched invoices with a difference above $10", int(var_rev), "Review"))
    other = (m["Oracle Bill-to"] != hdr["account"]).sum()
    if other:
        actions.append(("Confirm the receipt can be applied to invoices on other Metro-group bill-tos", int(other),
                        "Receipt (Note column)"))
    if (fmt["Status"] == "WO").any():
        actions.append(("After posting the receipt, create write-offs in Oracle (Bad Debt)",
                        int((fmt["Status"] == "WO").sum()), "WriteOffs"))
    if not c.all_pass:
        actions.insert(0, ("Fix the failed control checks before anything else", len(c.failed()), "Control_Checks"))
    actions += [
        ("Copy the Upload template columns into the AutoInvoice template (as text) and import", len(upload), "Upload"),
        ("Create, post and AutoMatch the receipt; check Amount Due = 0", len(receipt), "Receipt"),
        (f"Confirm DISC basis ({ctx['disc_basis']}) and receipt method ('{hdr['method']}') with the team", 0,
         "Rules workbook > Parameters"),
    ]
    act = pd.DataFrame([{"#": i, "Action": a, "Lines": n_, "Where": w} for i, (a, n_, w) in enumerate(actions, 1)])

    return [
        ("title", f"Metro Cash Application Summary: {ctx['region']} ({ctx['region_code']}), control {ctx['control']}"),
        ("subtitle", f"Remittance {ctx['file']}  |  receipt date {ctx['receipt_date']:%Y-%m-%d}  |  "
                     f"run {meta['run_at']:%Y-%m-%d %H:%M}"),
        ("section", "1. Run information"),
        ("note", "Which payment was processed, with which files and settings."),
        ("kv", [("Customer / region", f"Metro {ctx['region']} (file code {ctx['region_code']})"),
                ("Control number", ctx["control"]), ("Vendor (as on file)", ctx["vendor"]),
                ("Remittance file", ctx["file"]), ("Receipt / transaction date", ctx["receipt_date"]),
                ("Open Item Report", f"{meta['oi_file']} ({meta['oi_rows']:,} rows; "
                                     f"{meta['group_rows']:,} on Metro-group bill-tos)"),
                ("Closed-transactions export", meta["closed"] or "not supplied (NF lines are provisional)"),
                ("Rules / parameters", meta["rules"]), ("Run at", meta["run_at"].strftime("%Y-%m-%d %H:%M"))]),
        ("section", "2. Remittance reconciliation"),
        ("note", "Proves the CSV was read completely: gross - discount = net, and net = the total Metro paid. "
                 "Difference must be 0.00."),
        ("kv", [("Remittance lines", len(lines)), ("Gross", round(lines["Gross"].sum(), 2)),
                ("Discount (0.25% on invoices)", round(lines["Discount"].sum(), 2)),
                ("Net = cash paid", round(lines["Net"].sum(), 2)), ("Remittance summary total", total),
                ("Difference", round(lines["Net"].sum() - total, 2))]),
        ("section", "3. How every remittance line was classified"),
        ("note", "Each line gets one status. Matched lines go on the receipt; deductions, NF and SB are created in "
                 "Oracle through the upload (sign flipped); write-offs are absorbed."),
        ("note", "Line details: Formatted sheet (all lines), Deductions sheet, NF_SB sheet."),
        ("table", classes, {"status_col": "Status", "total": ["Lines", "% of lines", "Gross", "Discount",
                                                              "Net (paid)", "Upload amount", "On receipt"]}),
        ("section", "4. How the cash is explained"),
        ("note", "Receipt view: what is applied in Oracle (matched invoices + the NAC credit + write-offs) must add up "
                 "to the cash received. Difference must be 0.00."),
        ("table", receipt_bridge),
        ("note", "Upload view: the NAC line is the sum of all new transactions with the sign flipped, so the upload "
                 "totals 0. Upload lines: Upload sheet."),
        ("table", nac_build),
        ("section", "5. Receipt summary"),
        ("note", "The cash receipt to create in Oracle. Reference lines: Receipt sheet. The sum of reference amounts "
                 "plus write-offs must equal the receipt amount."),
        ("kv", [("Receipt number", hdr["number"]), ("Receipt amount", hdr["amount"]), ("Receipt date", hdr["date"]),
                ("Receipt method", hdr["method"]), ("Customer account", hdr["account"]),
                ("Batch name", hdr["batch"]), ("Reference lines", len(receipt)),
                ("Sum of reference amounts", round(receipt["Reference Amount"].sum(), 2)),
                ("Write-offs to create after posting", wo_total),
                ("Difference to receipt amount", round(receipt["Reference Amount"].sum() + wo_total - hdr["amount"],
                                                       2))]),
        ("section", "6. Action items"),
        ("note", "What a person has to do for this payment, in order. 'Where' is the sheet with the details."),
        ("table", act, {"wrap": True}),
    ]
