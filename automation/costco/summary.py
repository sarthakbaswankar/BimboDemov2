"""Detailed Summary sheet for a Costco run (sectioned layout, see excel_writer._write_blocks).

Line-level details live in their own sheets (regional sheets, Resolution_Log, Upload,
Receipt, WriteOffs); the control checks have their own Control_Checks sheet.
"""
import pandas as pd


REGIONS = ["Atlantic", "Western", "Ontario", "Quebec"]
ORDER = ["Matched", "Resolved", "WO", "SP", "OP", "NF", "Deduction"]
MEANING = {
    "Matched": "Found in Oracle, paid = open: applied on the receipt",
    "Resolved": "Difference explained by the engine (credit memo / neighbour / combination)",
    "WO": "Remaining difference within +/-$11: written off, not uploaded",
    "SP": "Costco paid less (> $11): short pay uploaded as debit",
    "OP": "Costco paid more (> $11): overpay uploaded as credit",
    "NF": "Invoice not in Oracle: uploaded as NF- at x -1",
    "Deduction": "RV / 65 / 14 / 604 / 13+ digit deduction: Ontario, uploaded at x -1",
}
RULE_NAMES = {
    "C-11": "C-11 Credit memo (same store, day, amount)",
    "C-12": "C-12 Neighbour invoice number",
    "C-13": "C-13 Combination of open items (same store)",
    "C-14": "C-14 Partial match (residual written off)",
}


def _docs(ln):
    return ([ln["primary"]] if ln["primary"] else []) + ln["added"]


def build_costco_summary(ctx):
    adv, lines, upload, receipt = ctx["adv"], ctx["lines"], ctx["upload"], ctx["receipt"]
    c, review, meta, hdr = ctx["controls"], ctx["review"], ctx["meta"], ctx["header"]
    region_billto, nac = ctx["region_billto"], ctx["nac_amounts"]
    total = adv.total
    n = len(lines)
    combined = [ln for ln in lines if len(ln["parts"]) > 1]

    # ---- 3 status overview -------------------------------------------------------------
    rows = []
    for s in ORDER:
        g = [ln for ln in lines if ln["status"] == s]
        if not g:
            continue
        rows.append({"Status": s, "Meaning": MEANING[s], "Lines": len(g), "% of lines": len(g) / n,
                     "Paid (remittance)": round(sum(ln["paid"] for ln in g), 2),
                     "Applied on receipt": round(sum(d["due"] for ln in g for d in _docs(ln)), 2),
                     "Upload amount": round(sum(ln["upload"] or 0 for ln in g), 2)})
    status_df = pd.DataFrame(rows)

    # ---- 4 / 5 regions -----------------------------------------------------------------
    matrix, fin = [], []
    for region in REGIONS:
        rl = [ln for ln in lines if ln["region"] == region]
        row = {"Region": region, "Bill-to": region_billto[region]}
        for s in ORDER:
            row[s] = sum(1 for ln in rl if ln["status"] == s)
        row["Total lines"] = len(rl)
        row["Paid"] = round(sum(ln["paid"] for ln in rl), 2)
        matrix.append(row)
        nac_ref, nac_amt = nac.get(region, ("-", 0.0))
        up_lines = [ln for ln in rl if ln["upload"] is not None]
        fin.append({"Region": region, "Bill-to": region_billto[region], "Paid": row["Paid"],
                    "Applied on receipt": round(sum(d["due"] for ln in rl for d in _docs(ln)), 2),
                    "Upload lines": len(up_lines), "Upload (excl. NAC)": round(sum(ln["upload"] for ln in up_lines), 2),
                    "NAC reference": nac_ref, "NAC amount": round(nac_amt, 2),
                    "Write-offs": round(sum(ln["wo"] or 0 for ln in rl if ln["status"] == "WO"), 2)})
    matrix_df, fin_df = pd.DataFrame(matrix), pd.DataFrame(fin)

    # ---- 6 auto-resolution ---------------------------------------------------------------
    worked = [ln for ln in lines if ln["status"] not in ("Matched", "Deduction")]
    outcome = {s: sum(1 for ln in worked if ln["status"] == s) for s in ["Resolved", "WO", "SP", "OP", "NF"]}
    log = ctx["log"]
    if len(log):
        log = log.assign(Rule=log["Rule"].str.extract(r"^(C-\d+)")[0].map(RULE_NAMES))
        by_rule = (log.groupby("Rule").agg(**{"Documents attached": ("Doc Trx", "size"),
                                              "Payment lines": ("Invoice", "nunique")}).reset_index())
    else:
        by_rule = pd.DataFrame()
    n_found_diff = sum(1 for ln in lines if ln["primary"] and abs(ln["var0"]) >= 0.005)
    n_na_inv = sum(1 for ln in lines if not ln["primary"] and ln["status"] != "Deduction")
    resolved_pct = (outcome["Resolved"] / len(worked)) if worked else 0.0

    # ---- 7 SP / OP -------------------------------------------------------------------------
    spop = pd.DataFrame([{
        "Region": ln["region"], "Invoice": ln["invoice"], "Store": ln["store"], "Invoice date": ln["date"],
        "Paid": ln["paid"], "Open in Oracle": round(sum(d["due"] for d in _docs(ln)), 2), "Net Var (L)": ln["L"],
        "Status": ln["status"], "Reference": ln["ref"],
        "Review": "Yes" if any("Large" in x for x in ln["notes"]) else ""}
        for ln in sorted(lines, key=lambda x: (x["region"], x["L"] or 0)) if ln["status"] in ("SP", "OP")])

    wo_total = ctx["wo_total"]
    n_wo = sum(1 for ln in lines if ln["status"] == "WO")
    bad12 = [r for r in receipt["Reference Value"] if r.isdigit() and len(r) != 12]

    # ---- 9 actions ------------------------------------------------------------------------
    def count_notes(word):
        return sum(1 for ln in lines if any(word in x for x in ln["notes"]))
    actions = []
    if not c.all_pass:
        actions.append(("Fix the failed control checks before anything else", len(c.failed()), "Control_Checks"))
    for word, text in [("Large", "Research large SP/OP lines (missing invoice or credit?) before upload"),
                       ("Partial", "Verify partial matches (C-14)"),
                       ("not in Oracle; documents found", "Verify engine matches on invoice numbers not in Oracle"),
                       ("More than one combination", "Verify lines where several combinations fit"),
                       ("belongs to", "Check region of lines where store and number pattern disagree"),
                       ("different dates", "Check combined lines with different dates")]:
        k = count_notes(word)
        if k:
            actions.append((text, k, "Review"))
    if adv.exceptions:
        actions.append(("Check PDF lines that could not be read", len(adv.exceptions), "Review"))
    if n_wo:
        actions.append(("After posting the receipt, create write-offs in Oracle (Bad Debt)", n_wo, "WriteOffs"))
    actions += [
        ("Copy the Upload template columns into the AutoInvoice template (as text) and import; fix duplicate rejects "
         "by adding the date", len(upload), "Upload"),
        ("Create, post and AutoMatch the receipt; check Amount Due = 0", len(receipt), "Receipt"),
        ("Open questions for the team: store 152 (Q17), payment terms and receipt account (Q18)", 0,
         "Rules workbook > Open_Questions"),
    ]
    act = pd.DataFrame([{"#": i, "Action": a, "Lines": k, "Where": w} for i, (a, k, w) in enumerate(actions, 1)])

    return [
        ("title", f"Costco Cash Application Summary: payment {adv.payment}"),
        ("subtitle", f"Payment advice {adv.path.name}  |  advice date {adv.advice_date:%Y-%m-%d}  |  work date "
                     f"{ctx['work_date']:%Y-%m-%d}  |  run {meta['run_at']:%Y-%m-%d %H:%M}"),
        ("section", "1. Run information"),
        ("note", "Which payment was processed, with which files and settings."),
        ("kv", [("Customer", "Costco (all regions, one payment)"), ("Payment #", adv.payment),
                ("Vendor #", adv.vendor), ("Currency", adv.currency), ("Payment advice file", adv.path.name),
                ("Pages", adv.pages), ("Advice date", adv.advice_date), ("Work / transaction date", ctx["work_date"]),
                ("Open Item Report", f"{meta['oi_file']} ({meta['oi_rows']:,} rows)"),
                ("Rules / parameters", meta["rules"]), ("Run at", meta["run_at"].strftime("%Y-%m-%d %H:%M"))]),
        ("section", "2. Remittance reconciliation"),
        ("note", "Proves the PDF was read completely: the lines read must add up exactly to the Remittance TOTAL "
                 "printed by Costco. Difference must be 0.00."),
        ("kv", [("PDF lines read", len(ctx["raw"])), ("Sum of PDF lines", round(ctx["raw"]["Amount"].sum(), 2)),
                ("Remittance TOTAL (printed)", total),
                ("Difference", round(ctx["raw"]["Amount"].sum() - (total or 0), 2)),
                ("Invoice numbers listed more than once (combined)",
                 f"{len(combined)} invoices from {sum(len(x['parts']) for x in combined)} lines"),
                ("Payment lines after combining", n), ("Unreadable lines", len(adv.exceptions))]),
        ("section", "3. How every payment line was classified"),
        ("note", "Each payment line gets one status. Matched and Resolved lines go on the receipt; SP, OP, NF and "
                 "deductions are created in Oracle through the upload; write-offs are absorbed."),
        ("note", "Line details: the Atlantic / Western / Ontario / Quebec sheets; write-offs on the WriteOffs sheet."),
        ("table", status_df, {"status_col": "Status", "total": ["Lines", "% of lines", "Paid (remittance)",
                                                                "Applied on receipt", "Upload amount"]}),
        ("section", "4. Regions: line count by status"),
        ("note", "How many payment lines of each status fall in each region. The region comes from the bill-to "
                 "in Oracle; all deductions go to Ontario."),
        ("table", matrix_df, {"total": ORDER + ["Total lines", "Paid"]}),
        ("section", "5. Regions: money"),
        ("note", "For each region: Paid = Applied on receipt + NAC amount + Write-offs. The NAC line balances that "
                 "region's upload lines (SP, OP, NF, deductions) to 0."),
        ("table", fin_df, {"total": ["Paid", "Applied on receipt", "Upload lines", "Upload (excl. NAC)", "NAC amount",
                                     "Write-offs"]}),
        ("section", "6. Automatic resolution of differences"),
        ("note", "Lines where paid and open did not agree, or the invoice number was not in Oracle. The engine tried "
                 "credit memos, neighbour invoice numbers (x+1, x+2, x-1) and combinations for the same store and day. "
                 "Every attached document is listed on the Resolution_Log sheet."),
        ("kv", [("Found in Oracle with a difference", n_found_diff),
                ("Invoice numbers not in Oracle (searched by store and date)", n_na_inv),
                ("Lines worked by the engine", len(worked)),
                ("Fully resolved (difference 0)", outcome["Resolved"]),
                ("Ended as write-off", outcome["WO"]), ("Ended as short pay", outcome["SP"]),
                ("Ended as overpay", outcome["OP"]), ("Ended as NF", outcome["NF"]),
                ("Share fully resolved", f"{resolved_pct:.0%}")]),
        ("note", "Documents attached by rule (a line can use more than one rule, so payment lines are not totalled):"),
        ("table", by_rule, {"total": ["Documents attached"]}),
        ("section", "7. Short pays and overpays"),
        ("note", "Differences above $11 that nothing in Oracle explains. SP = Costco paid less (debit), OP = Costco paid "
                 "more (credit). 'Review = Yes' means above $1,000: check for a missing invoice or credit first."),
        ("table", spop, {"status_col": "Status", "total": ["Paid", "Open in Oracle", "Net Var (L)"]}),
        ("section", "8. Receipt summary"),
        ("note", "The cash receipt to create in Oracle. Reference lines: Receipt sheet. The sum of reference amounts "
                 "plus write-offs must equal the receipt amount."),
        ("kv", [("Receipt number", hdr["number"]), ("Receipt amount", hdr["amount"]), ("Receipt date", hdr["date"]),
                ("Receipt method", hdr["method"]), ("Customer account", hdr["account"]), ("Batch name", hdr["batch"]),
                ("Reference lines", len(receipt)),
                ("Sum of reference amounts", round(receipt["Reference Amount"].sum(), 2)),
                ("Write-offs to create after posting", wo_total),
                ("Difference to receipt amount",
                 round(receipt["Reference Amount"].sum() + wo_total - hdr["amount"], 2)),
                ("References not 12 digits (deductions, OK)", ", ".join(bad12) or "none")]),
        ("section", "9. Action items"),
        ("note", "What a person has to do for this payment, in order. 'Where' is the sheet with the details."),
        ("table", act, {"wrap": True}),
    ]
