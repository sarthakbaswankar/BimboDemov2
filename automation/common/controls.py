"""PASS/FAIL control checks shown on the Control_Checks sheet."""
import pandas as pd

# (control name starts with, what it checks, what to do if it fails)
CONTROL_HELP = [
    ("Remittance lines net = summary total",
     "The net amounts of all invoice lines read from the Metro CSV add up to the total in the REMITTANCE SUMMARY row.",
     "A line was missed or misread. Check the CSV (and the Other_Records sheet, if present) and run again. "
     "Do not upload."),
    ("Gross - discount = net",
     "For the whole file, gross minus discount equals net, so the 0.25% discount is consistent with what Metro paid.",
     "The file contains an inconsistent line. Find it on Raw_Remittance and ask Metro / the team."),
    ("PDF lines = Remittance TOTAL",
     "The lines read from the Costco PDF add up exactly to the Remittance TOTAL printed on the last page.",
     "The PDF layout may have changed or a line could not be read (see 'parse exception' rows on Review). "
     "Do not upload."),
    ("Combined lines total",
     "After adding together invoice numbers that appear more than once, the total is unchanged.",
     "Report to the automation owner. Do not upload."),
    ("Every line has exactly one status",
     "Every line was classified exactly once (Matched, Resolved, Deduction, NF, SB, SP, OP, WO ...). Nothing was lost "
     "or counted twice.", "Report to the automation owner. Do not upload."),
    ("Upload",
     "The new transactions to create plus the NAC balancing line add up to exactly 0 (for Costco: per region), so the "
     "clearing account stays at 0.",
     "A line is missing from the upload or has the wrong sign. Do not upload; check the Upload sheet."),
    ("Receipt references + write-offs = remittance total",
     "Everything applied on the receipt (matched documents + NAC) plus the write-offs equals the money received.",
     "The receipt will not balance in Oracle (Amount Due not 0). Check the Review lines and variances before "
     "creating the receipt."),
    ("No open item",
     "No Oracle invoice or credit memo is applied to more than one payment line.",
     "The same document would be paid twice. Do not upload; report to the automation owner."),
    ("Receipt invoice numbers are 12 digits",
     "Every invoice number on the receipt has 12 digits (a leading 0 is restored where the export dropped it).",
     "Oracle will not find the invoice. Correct the number on the Receipt sheet (keep it as text)."),
    ("Write-offs within",
     "Every write-off is within the customer tolerance (Metro +/-$10, Costco +/-$11, limits included).",
     "A larger difference was treated as a write-off. Report to the automation owner."),
    ("No duplicate upload reference",
     "Every transaction number in the upload is used only once.",
     "AutoInvoice would reject the duplicate. Check the references (suffix -1 / -2) on the Upload sheet."),
]


def describe(name):
    for prefix, what, if_fail in CONTROL_HELP:
        if name.startswith(prefix):
            return what, if_fail
    return "", ""


class Controls:
    def __init__(self):
        self.rows = []

    def amount(self, name, expected, actual, tol=0.01, note=""):
        diff = round(actual - expected, 2)
        ok = abs(diff) <= tol
        self.rows.append({"Control": name, "Expected": round(expected, 2), "Actual": round(actual, 2),
                          "Difference": diff, "Result": "PASS" if ok else "FAIL", "Note": note})
        return ok

    def check(self, name, ok, note=""):
        self.rows.append({"Control": name, "Expected": None, "Actual": None, "Difference": None,
                          "Result": "PASS" if ok else "FAIL", "Note": note})
        return ok

    @property
    def all_pass(self):
        return all(r["Result"] == "PASS" for r in self.rows)

    def failed(self):
        return [r["Control"] for r in self.rows if r["Result"] != "PASS"]

    def frame(self):
        return pd.DataFrame(self.rows, columns=["Control", "Expected", "Actual", "Difference", "Result", "Note"])

    def described_frame(self):
        rows = []
        for i, r in enumerate(self.rows, start=1):
            what, if_fail = describe(r["Control"])
            rows.append({"#": i, "Control": r["Control"], "What it checks": what, "Result": r["Result"],
                         "Expected": r["Expected"], "Actual": r["Actual"], "Difference": r["Difference"],
                         "Detail": r["Note"], "What to do if it fails": if_fail})
        return pd.DataFrame(rows)

    def sheet_blocks(self, title, subtitle):
        """Sectioned Control_Checks sheet (see excel_writer._write_blocks)."""
        n, failed = len(self.rows), self.failed()
        banner = ((f"ALL {n} CHECKS PASS: the totals tie and the upload and receipt balance.", "READY")
                  if not failed else
                  (f"{len(failed)} OF {n} CHECKS FAILED: do not upload until fixed ({'; '.join(failed)}).", "BLOCKED"))
        return [
            ("title", title),
            ("subtitle", subtitle),
            ("banner", *banner),
            ("section", "What this sheet is"),
            ("note", "Automatic checks run on every file. They prove the file was read completely, every line was "
                     "classified once, and the upload and receipt balance."),
            ("note", "All checks must say PASS before anything is uploaded to Oracle. "
                     "For a FAIL, follow the 'If it fails' column."),
            ("note", "Amounts: Expected = the figure the file or the rules say it should be, Actual = what the "
                     "automation calculated, Difference = Actual - Expected (up to 0.01 counts as equal)."),
            ("section", "Check results"),
            ("table", self.described_frame(), {"status_col": "Result", "wrap": True}),
        ]
