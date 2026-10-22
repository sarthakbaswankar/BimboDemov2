"""Upload (AutoInvoice / FBDI RA_INTERFACE_LINES_ALL layout) and receipt sheet builders."""
import pandas as pd

UPLOAD_COLUMNS = [
    "Business Unit Name", "Transaction Batch Source Name", "Transaction Type Name", "Payment Terms",
    "Transaction Date", "Accounting Date", "Transaction Number", "Bill-to Customer Account Number",
    "Bill-to Customer Site Number", "Transaction Line Type", "Transaction Line Description",
    "Currency Code", "Transaction Line Amount", "Line Transactions Flexfield Segment 1",
    # helper columns (not part of the FBDI file)
    "Category", "Region", "NAC Group", "Source Invoice",
]

RECEIPT_COLUMNS = [
    "Receipt Number", "Receipt Amount", "Receipt Date", "Receipt Method", "Customer Account Number",
    "Batch Name", "Sequence Number", "Reference Type", "Reference Value", "Reference Amount",
    "Region", "Bill-to of Document", "Category", "Note",
]


def upload_frame(lines, common, terms_dr="", terms_cr=""):
    """lines: dicts with ref, amount, bill_to, trx_date, category, region, group, source."""
    rows = []
    for ln in lines:
        amt = round(ln["amount"], 2)
        debit = amt > 0
        rows.append({
            "Business Unit Name": common["BUSINESS_UNIT"],
            "Transaction Batch Source Name": common["BATCH_SOURCE"],
            "Transaction Type Name": common["TRX_TYPE_DR"] if debit else common["TRX_TYPE_CR"],
            "Payment Terms": terms_dr if debit else terms_cr,
            "Transaction Date": ln["trx_date"].strftime("%Y-%m-%d"),
            "Accounting Date": ln["trx_date"].strftime("%Y-%m-%d"),
            "Transaction Number": ln["ref"],
            "Bill-to Customer Account Number": ln["bill_to"],
            "Bill-to Customer Site Number": ln["bill_to"],
            "Transaction Line Type": common["LINE_TYPE"],
            "Transaction Line Description": ln["ref"],
            "Currency Code": common["CURRENCY"],
            "Transaction Line Amount": amt,
            "Line Transactions Flexfield Segment 1": ln["ref"],
            "Category": ln.get("category", ""),
            "Region": ln.get("region", ""),
            "NAC Group": ln.get("group", ""),
            "Source Invoice": ln.get("source", ""),
        })
    df = pd.DataFrame(rows, columns=UPLOAD_COLUMNS)
    # lowest (largest credit) first, highest (largest debit) last
    return df.sort_values("Transaction Line Amount", kind="stable").reset_index(drop=True)


def receipt_frame(header, refs):
    """header: receipt-level fields; refs: dicts with value, amount, region, bill_to, category, note."""
    rows = []
    for r in sorted(refs, key=lambda x: x["amount"]):
        rows.append({
            "Receipt Number": header["number"],
            "Receipt Amount": round(header["amount"], 2),
            "Receipt Date": header["date"].strftime("%Y-%m-%d"),
            "Receipt Method": header["method"],
            "Customer Account Number": header["account"],
            "Batch Name": header.get("batch", ""),
            "Sequence Number": 1,
            "Reference Type": "Transaction Number",
            "Reference Value": r["value"],
            "Reference Amount": round(r["amount"], 2),
            "Region": r.get("region", ""),
            "Bill-to of Document": r.get("bill_to", ""),
            "Category": r.get("category", ""),
            "Note": r.get("note", ""),
        })
    return pd.DataFrame(rows, columns=RECEIPT_COLUMNS)
