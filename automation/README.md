# Metro & Costco cash application automation

Turns a customer remittance plus the Oracle Open Item Report into one Excel workbook per run. The workbook covers everything up to the AutoInvoice (FBDI) and receipt upload. Generating the FBDI file itself is out of scope for now.

## Virtual environment
The main code has its own Python environment in `CODES\.venv`, built from `CODES\requirements.txt` (the TESTING copy has a separate one). Activate it once per terminal from the `CODES` folder with `.venv\Scripts\Activate.ps1` (PowerShell) or `.venv\Scripts\activate.bat` (cmd), or call `.venv\Scripts\python.exe` directly.
To rebuild it: `python -m venv .venv` then `.venv\Scripts\python.exe -m pip install -r requirements.txt`.

## Run (from the `CODES` folder)

```
python automation/run_metro.py  --remit "CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv" --open-items "Open Item Reports/Open Item report Metro.xlsx"
python automation/run_costco.py --remit "Costco Payment.PDF" --open-items "Open Item Reports/open item report costco.xlsx"
```

| Option | Customer | Purpose |
|---|---|---|
| `--closed-trx <xlsx/csv>` | Metro | Export of closed/all transactions. Lets the engine decide NF vs SB instead of marking NF as provisional |
| `--receipt-date YYYY-MM-DD` | Metro | Defaults to the date in the file name |
| `--region Ontario/Quebec` | Metro | Only needed if the file name has no CA5/CA6 |
| `--work-date YYYY-MM-DD` | Costco | Transaction date. Defaults to the advice date |
| `--force` | Costco | Write the workbook even if the PDF lines don't tie to the Remittance TOTAL |
| `--out <folder>` | both | Defaults to `CODES/Output` |

Output files: `Output/Metro_<Region>_<Control>_<Date>.xlsx` and `Output/Costco_<Payment>_<Date>.xlsx`.

Each workbook opens with two sheets:
- **Summary**: a status banner, then run information, reconciliation, line classification, the cash bridge, the receipt summary and action items. Costco adds regions, automatic resolution and short pays / overpays. Every section has a one-line description. Line-level details stay in their own sheets.
- **Control_Checks**: every automatic check, with what it checks, PASS/FAIL, the amounts and what to do if it fails.

To get a printable copy as HTML and PDF (uses Microsoft Edge; Excel is not needed; add `--sheet Control_Checks` for the checks):

```
python automation/preview_summary.py Output/Metro_Ontario_2000238539_20260827.xlsx Output/Costco_1509518_20260901.xlsx
```
This writes `Output/Summary_Sample_Metro.pdf/.html` and `Output/Summary_Sample_Costco.pdf/.html`.

## Daily checklist
1. Run the Open Item Report the same morning, before processing.
2. Run the script.
3. Open the workbook. The Summary banner must not be red, and **Control_Checks** must be all PASS.
4. Work through the **Review** sheet: provisional NFs, engine matches on invoices not found in Oracle, partial matches and large SP/OP lines.
5. Use **Upload** (AutoInvoice lines) and **Receipt** (receipt references). **WriteOffs** lists the amounts to write off in Oracle after posting.

## Rules and parameters
`Set of rules for automation/National_Customers_Automation_Rules.xlsx` documents every rule (rule IDs such as M-05 and C-12 appear in each line's *Rule* column). Its **Parameters** sheet is read on every run, so you can change tolerances, bill-tos, transaction types and terms there without editing code. Run `python automation/build_rules_workbook.py` to regenerate the workbook; this resets Parameters to the defaults in `config.py`.

`Set of rules for automation/National_Customers_Rules_Guide.xlsx` is the plain-language version of the rules, for new team members and reviewers. It covers the glossary, process overview, decision flow, each rule explained with examples, worked examples, an output guide and a daily checklist. The code does not read it. Regenerate it with `python automation/build_rules_guide.py`.

## Code layout
| Path | Role |
|---|---|
| `config.py` | Parameter defaults + loader |
| `common/` | Open item loader, key normalisation, reference naming/suffixing, upload/receipt builders, controls, Excel writer |
| `metro/` | CSV parser + engine |
| `costco/` | PDF parser, engine, variance resolution engine (`resolver.py`) |
| `../tests/` | Regression tests on the sample files: `python -m pytest tests -q` |
