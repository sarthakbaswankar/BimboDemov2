# demo National Customers: Cash Application API

A FastAPI service around the Metro / Costco cash application automation. The processing code in `automation\` is a copy of the main `CODES\automation` (same matching and calculation logic). The open item reports in `Open Item Reports\` are bundled unchanged.

Upload a remittance and get the output workbook the automation produces: Summary, Raw_Remittance, Formatted, Upload, Receipt, WriteOffs, then the other sheets, and Control_Checks as the last sheet (same order as CODES and TESTING).

Differences from the main `CODES\automation` (22-Oct-2026):
- **Output colour scheme** (`common\excel_writer.py`): rows stay white with thin row lines and negatives are black, so colour always carries meaning.
  - Navy header rows and navy section bars on the Summary.
  - The Status cell only carries Excel's standard Good / Neutral / Bad colours: green = Matched / Resolved / PASS; amber = WO / SP / OP; red = NF / SB / Review / FAIL; blue = Deduction / DISC / NAC.
  - Documents attached to a Costco payment line are in grey italics.
  - The Summary has no result banner; it starts with the title and section 1. The Control_Checks sheet keeps its banner (green / amber / red).
- **Control_Checks is the last sheet** (`metro\engine.py`, `costco\engine.py`); the X-Controls response header also gives the result.
- **Metro CSV reader from TESTING** (`metro\parser.py`): columns are found by their header name, and names like `METRO_ON 08-27-2026 $164986.02 CA6.csv` are accepted (CA5/CA6 and the date are read from the name).

## Folder
| Path | What |
|---|---|
| `api\index.py` | The API (Vercel entry point) |
| `automation\` | The automation code (same as CODES) |
| `Open Item Reports\` | Open item reports used by the API (Metro, Costco) |
| `Set of rules for automation\` | Rules workbook (its Parameters sheet is read on every run) |
| `tests\` | `test_regression.py` (same as CODES) + `test_api.py` |
| `vercel.json`, `.vercelignore`, `requirements.txt` | Vercel deployment |
| `requirements-dev.txt` | Local extras (uvicorn, httpx, pytest) |
| sample `CAN_RMT_…csv`, `Costco Payment.PDF` | Test inputs (not deployed) |

## API
**`POST /process_national_customers/upload`** (multipart/form-data)

| Field | Required | Value |
|---|---|---|
| `customer` | yes | `metro` or `costco`, in any letter case (`Metro`, `METRO`, `Costco`, …) |
| `filename` | yes | File name, e.g. `CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv` or `Costco Payment.PDF`; the extension may be left out. Metro: as in the main automation, `CA6` = Ontario and `CA5` = Quebec, and the receipt date is read from the name. If this name has no CA5/CA6, the uploaded file's own name is used. Costco: any name; everything is read from the PDF |
| `file` | yes | Metro `.csv` / Costco `.pdf` (max 4 MB) |
| `output` | no | `excel` (default: returns the workbook) or `json` (returns a short summary) |

The Open Item Reports are **not** sent with the request. The two reports in `Open Item Reports\` are read once when the service starts (about 5 s) and kept in memory; every request is matched against them (`common\open_items.py`). Each run gets its own copy, so runs don't affect each other. If a report file is replaced, it is read again on the next request. `GET /` shows how many rows of each report are in memory.

The response is the output workbook (`Metro_<Region>_<Control>_<Date>.xlsx` / `Costco_<Payment>_<Date>.xlsx`). The header `X-Controls` says `ALL PASS` or which checks failed.

**`POST /process_national_customers`** (application/json): the same input as JSON, with the file as base64 text. The response is the same.
```json
{
  "customer": "Costco",
  "filename": "Costco Payment.PDF",
  "file_base64": "<base64 of the file>",
  "output": "base64"
}
```
`python make_payload.py costco "Costco Payment.PDF"` writes a ready payload to `payloads\`. Base64 makes the body about a third larger, so on Vercel (4.5 MB limit) a JSON payload can carry a file of up to about 3 MB.

`output` values:
- `excel` (default): the workbook as a download.
- `json`: summary only.
- `base64`: summary plus the workbook inside the JSON (`output_file_name`, `content_type`, `file_base64`). Use this for **Oracle Visual Builder** and other REST clients that expect JSON.

Every response has `"status": "success"` or `"status": "error"`. Errors look like `{"status": "error", "code": 400, "detail": "..."}`.

**Browser apps (Oracle Visual Builder):** CORS is on. Set the Vercel environment variable `CORS_ORIGINS` to your VB instance URL (e.g. `https://<instance>.oraclecloud.com`) to allow only that app; the default allows any origin.

The file type is checked from the file's content, not its name. Costco must be a PDF; Metro must be a text CSV with REMITTANCE records.

Errors:
- `400`: unknown customer, or the file doesn't match the customer (e.g. a CSV sent as Costco)
- `413`: file too large
- `422`: file can't be processed, e.g. Metro with no CA5/CA6 in either name, or the Costco PDF total doesn't tie
- `401`: wrong API key (see below)

Other endpoints:
- `GET /`: service info and which open item reports are bundled
- `GET /health`
- `GET /docs`: interactive page to try the API in the browser

## Run locally (Windows, cmd)
```bat
cd "C:\Users\Sarthak Basawankar\OneDrive - acsesolutions.com\Desktop\Bimbo National Customers\CODES\demo National Customers"
.venv\Scripts\activate.bat
uvicorn api.index:app --reload --port 8000
```
Open http://127.0.0.1:8000/docs, choose **POST /process_national_customers/upload** and click **Try it out**.

Command-line example:
```bat
curl -o metro_output.xlsx -F "customer=metro" -F "filename=CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv" -F "file=@CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv" http://127.0.0.1:8000/process_national_customers/upload
curl -o costco_output.xlsx -F "customer=costco" -F "filename=Costco Payment.PDF" -F "file=@Costco Payment.PDF" http://127.0.0.1:8000/process_national_customers/upload
```
Tests: `python -m pytest tests -q` (39 tests; they also check the JSON payload, the base64 response, CORS, error format, file names handled like the main automation, any-case customer names, the in-memory reports, and that Summary is the first sheet and Control_Checks the last).

To rebuild the environment: `python -m venv .venv`, then `.venv\Scripts\python.exe -m pip install -r requirements-dev.txt`.

## Deploy to Vercel
1. Install the Vercel CLI (needs Node.js): `npm i -g vercel`, then `vercel login`.
2. From this folder: `vercel` (first time: create a new project), then `vercel --prod`.
   Or push this folder to a Git repository and import it at vercel.com → **Add New → Project**.
3. **Protect it:** set an environment variable `API_KEY` in the Vercel project (Settings → Environment Variables), then redeploy. Every request must then send the header `X-API-Key: <your key>`, e.g. `curl -H "X-API-Key: <key>" …`. You can also turn on Vercel Deployment Protection.
4. Test it: `https://<your-project>.vercel.app/health` and `/docs`.

**Vercel limits to keep in mind:**
- Request body up to 4.5 MB.
- The function may run up to 60 s; it's set in `vercel.json`. Locally, start-up (both reports into memory) takes about 5 s, then a request takes about 5–6 s. On Vercel, the first request after an idle period also loads the reports.
- Files are only written to `/tmp`, and are removed after each request.

**Data:** the bundled open item reports contain customer financial data, and they become part of the deployment. Only deploy to a private, protected project.

**Updating the open item reports:** replace the files in `Open Item Reports\` (keep the same names) and redeploy. Locally, the API picks up a replaced file on the next request.
