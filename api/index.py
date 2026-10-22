"""FastAPI service for the Metro / Costco cash application automation (Vercel entry point).

POST /process_national_customers/upload  (multipart/form-data)
    customer    metro | costco (any letter case: Metro, METRO, costco, ...)
    filename    original file name, e.g. CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv or Costco Payment.PDF
                (Metro: the region CA5/CA6 and the receipt date are read from this name)
    file        the remittance file (Metro .csv, Costco .pdf)
    output      optional: excel (default, returns the output workbook) | json (returns a short summary)

POST /process_national_customers  (application/json) - the same, with the file as base64 text:
    {"customer": "Costco", "filename": "Costco Payment.PDF", "file_base64": "<base64>", "output": "base64"}   (default base64 = JSON back)

The Open Item Reports are not sent with the request: the two reports in 'Open Item Reports/' are read once
when the service starts and kept in memory; every request is matched against them.
The processing is exactly the automation in automation/ (same as the CODES folder); this file only wraps it.
"""
import base64
import binascii
import os
from contextlib import asynccontextmanager
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import parse_qsl, urlencode

from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.exceptions import HTTPException as StarletteHTTPException

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "automation"))

from common.keys import metro_key, r12  # noqa: E402
from common.open_items import cached_reports, load_open_items  # noqa: E402
from costco.engine import run_costco  # noqa: E402
from metro.engine import run_metro  # noqa: E402
from metro.parser import _region_code  # noqa: E402

OPEN_ITEMS = {
    "metro": ROOT / "Open Item Reports" / "Open Item report Metro.xlsx",
    "costco": ROOT / "Open Item Reports" / "open item report costco.xlsx",
}
KEY_FUNCS = {"metro": metro_key, "costco": r12}
EXTENSIONS = {"metro": ".csv", "costco": ".pdf"}
MAX_UPLOAD_BYTES = 4 * 1024 * 1024  # Vercel limits a request body to 4.5 MB
# Running locally: every output workbook is also saved here (handy when the browser can't download,
# e.g. VS Code's built-in browser). Not on Vercel, where only /tmp is writable.
LOCAL_OUTPUT = None if os.environ.get("VERCEL") else ROOT / "Output"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def preload_open_items():
    """Read both Open Item Reports into memory (later requests reuse them)."""
    for cust, path in OPEN_ITEMS.items():
        if path.exists():
            load_open_items(path, KEY_FUNCS[cust])


@asynccontextmanager
async def lifespan(_app):
    try:
        preload_open_items()
    except Exception as exc:  # the first request will try again and report the error
        print(f"Open Item Reports not preloaded: {exc}")
    yield


app = FastAPI(
    title="Bimbo National Customers - Cash Application",
    version="1.1.0",
    description="Upload a Metro remittance (.csv) or a Costco payment advice (.pdf) and get the output workbook "
                "(Summary, Upload, Receipt, ..., Control_Checks). It is matched against the Open Item Reports "
                "held in memory. Same logic as the CODES automation.",
    lifespan=lifespan,
)

class VercelPathMiddleware:
    """Puts back the path the caller asked for.

    vercel.json rewrites every request to the function /api/index and passes the original path in the query
    parameter __vpath. Depending on the project preset, Vercel hands the app either the original path or the
    rewrite destination (/api/index); in the second case the original path is restored from __vpath.
    """

    def __init__(self, asgi_app):
        self.app = asgi_app

    async def __call__(self, scope, receive, send):
        if scope["type"] == "http":
            query = parse_qsl(scope.get("query_string", b"").decode("latin-1"), keep_blank_values=True)
            vpath = next((v for k, v in query if k == "__vpath"), None)
            if vpath is not None or scope["path"].startswith("/api/index"):
                scope = dict(scope)
                scope["query_string"] = urlencode([(k, v) for k, v in query if k != "__vpath"]).encode("latin-1")
                if scope["path"].startswith("/api/index"):
                    path = "/" + (vpath or "").lstrip("/")
                    scope["path"], scope["raw_path"] = path, path.encode("utf-8")
        await self.app(scope, receive, send)


app.add_middleware(VercelPathMiddleware)

# Browser apps (e.g. Oracle Visual Builder) call the API from another domain, so CORS must allow it.
# CORS_ORIGINS = comma-separated list of allowed origins, e.g. https://<instance>.oraclecloud.com (default: any).
app.add_middleware(
    CORSMiddleware,
    allow_origins=[o.strip() for o in os.environ.get("CORS_ORIGINS", "*").split(",") if o.strip()],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "X-API-Key"],
    expose_headers=["X-Controls", "X-Saved-To", "Content-Disposition"],
)


@app.exception_handler(StarletteHTTPException)
async def http_error(_request, exc):
    """Every error has the same JSON shape: {"status": "error", "code": 422, "detail": "..."}."""
    return JSONResponse({"status": "error", "code": exc.status_code, "detail": exc.detail},
                        status_code=exc.status_code, headers=getattr(exc, "headers", None))


@app.exception_handler(RequestValidationError)
async def validation_error(_request, exc):
    missing = ", ".join(".".join(str(x) for x in e["loc"][1:]) or "body" for e in exc.errors())
    return JSONResponse({"status": "error", "code": 422,
                         "detail": f"Request body is not valid; check these fields: {missing}"}, status_code=422)


def _base_name(value):
    return Path((value or "").strip().replace("\\", "/")).name


def work_file_name(cust, typed, uploaded):
    """Name the remittance is saved under, as the main automation would see it.

    Metro (like run_metro.py): the region CA5/CA6 and the receipt date come from the file name, so the typed
    filename is used if it has CA5/CA6, otherwise the uploaded file's own name. Costco: the PDF holds everything,
    any name will do. The right extension is added if it is missing ('Costco Payment' -> 'Costco Payment.pdf').
    """
    names = [n for n in (_base_name(typed), _base_name(uploaded)) if n]
    if cust == "metro":
        name = next((n for n in names if _region_code(n)), names[0] if names else "remittance")
    else:
        name = names[0] if names else "payment"
    ext = EXTENSIONS[cust]
    if not name.lower().endswith(ext):
        name = (Path(name).stem if Path(name).suffix.lower() in (".csv", ".pdf", ".txt") else name) + ext
    return name


def check_content(cust, data):
    """The file itself must be what the customer sends: Costco a PDF, Metro a remittance CSV (text)."""
    is_pdf = data[:1024].lstrip().startswith(b"%PDF")
    if cust == "costco" and not is_pdf:
        raise HTTPException(400, "Customer is Costco, but the uploaded file is not a PDF. Costco sends its payment "
                                 "advice as a PDF.")
    if cust == "metro":
        if is_pdf:
            raise HTTPException(400, "Customer is Metro, but the uploaded file is a PDF. Metro sends a .csv "
                                     "remittance (CAN_RMT_METRODE-CA5/CA6-...).")
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            raise HTTPException(400, "Customer is Metro, but the uploaded file is not a text (.csv) file.") from None
        if "REMITTANCE" not in text.upper():
            raise HTTPException(400, "This does not look like a Metro remittance CSV (no REMITTANCE records).")


def check_api_key(x_api_key):
    expected = os.environ.get("API_KEY")
    if expected and x_api_key != expected:
        raise HTTPException(status_code=401, detail="Missing or wrong X-API-Key header.")


@app.get("/")
def info():
    return {
        "service": "Bimbo National Customers - Cash Application",
        "customers": sorted(OPEN_ITEMS),
        "endpoints": {"POST /process_national_customers/upload": "form-data: customer, filename, file (+ optional output)",
                      "POST /process_national_customers": "JSON: customer, filename, file_base64 (+ optional output)",
                      "GET /health": "health check", "GET /docs": "interactive API documentation"},
        "open_item_reports": {c: {"file": p.name, "present": p.exists(),
                                  "size_kb": round(p.stat().st_size / 1024) if p.exists() else None,
                                  "rows_in_memory": cached_reports().get(p.name)}
                              for c, p in OPEN_ITEMS.items()},
        "api_key_required": bool(os.environ.get("API_KEY")),
    }


@app.get("/health")
def health():
    return {"status": "ok"}


class ProcessRequest(BaseModel):
    """JSON input for POST /process_national_customers: the same fields as the form, with the file as base64 text."""
    customer: str = Field(..., description="metro or costco, in any letter case")
    filename: str = Field(..., description="e.g. CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv or Costco Payment.PDF "
                                           "(Metro: CA6 = Ontario, CA5 = Quebec, read from this name)")
    file_base64: str = Field(..., description="The file's content, base64 encoded (a 'data:...;base64,' prefix is "
                                              "allowed)")
    output: str = Field("base64", description="base64 (default: JSON with the summary + the workbook as base64, "
                                              "for Oracle Visual Builder), json (summary only) or excel (the "
                                              "workbook as a file download)")

    model_config = {"json_schema_extra": {"examples": [
        {"customer": "Costco", "filename": "Costco Payment.PDF", "file_base64": "JVBERi0xLjQKJ...", "output": "base64"},
        {"customer": "Metro", "filename": "CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv",
         "file_base64": "UkVNSVRUQU5DRSBTVU1NQVJZ...", "output": "base64"},
    ]}}


@app.post("/process_national_customers/upload")
async def process_national_customers_upload(
    customer: str = Form(..., description="metro or costco, in any letter case (Metro, METRO, Costco, ...)"),
    filename: str = Form(..., description="File name, e.g. CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv or Costco "
                                          "Payment.PDF. Metro: CA6 = Ontario, CA5 = Quebec, read from this name "
                                          "or from the uploaded file's own name"),
    file: UploadFile = File(..., description="Metro remittance .csv or Costco payment advice .pdf"),
    output: str = Form("excel", description="excel (default), json or base64"),
    x_api_key: str | None = Header(None),
):
    """Form-data input: customer, filename, file (+ optional output)."""
    check_api_key(x_api_key)
    return run_request(customer, filename, file.filename, await file.read(), output)


@app.post("/process_national_customers")
def process_national_customers(body: ProcessRequest, x_api_key: str | None = Header(None)):
    """JSON input: {"customer", "filename", "file_base64", "output"}. Same processing and response as /process_national_customers/upload."""
    check_api_key(x_api_key)
    b64 = body.file_base64.strip()
    if b64.startswith("data:") and "," in b64:
        b64 = b64.split(",", 1)[1]
    try:
        data = base64.b64decode("".join(b64.split()), validate=True)
    except (binascii.Error, ValueError):
        raise HTTPException(400, "file_base64 is not valid base64.") from None
    return run_request(body.customer, body.filename, body.filename, data, body.output)


def run_request(customer, filename, uploaded_name, data, output):
    """Shared by both endpoints: checks the input, runs the automation and builds the response."""
    cust = customer.strip().casefold()
    if cust not in OPEN_ITEMS:
        raise HTTPException(400, f"Unknown customer '{customer}'. Use one of: {', '.join(sorted(OPEN_ITEMS))}.")
    if not data:
        raise HTTPException(400, "The uploaded file is empty.")
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "The uploaded file is larger than 4 MB.")
    check_content(cust, data)
    name = work_file_name(cust, filename, uploaded_name)
    if cust == "metro" and not _region_code(name):    # Metro only; Costco has no region in the name
        raise HTTPException(422, "Ontario or Quebec? Neither the filename nor the uploaded file's name contains "
                                 "CA6 (Ontario) or CA5 (Quebec), e.g. CAN_RMT_METRODE-CA6-RMT_2252_202608271334.csv. "
                                 "Send the original Metro file name.")

    work = Path(tempfile.mkdtemp(prefix="national_customers_"))  # /tmp on Vercel (the only writable place)
    try:
        remit = work / name
        remit.write_bytes(data)
        oi_path = OPEN_ITEMS[cust]                 # bundled report, matched from memory
        if not oi_path.exists():
            raise HTTPException(500, f"Open Item Report not found: {oi_path.name}")
        out_dir = work / "output"
        out_dir.mkdir()

        if cust == "metro":
            results = run_metro(remit, oi_path, out_dir)        # region from CA5/CA6 in the name, as in main
            summary = [{"control": r["control"], "remittance_total": r["total"], "nac_reference": r["nac_ref"],
                        "nac_amount": r["nac"], "status_counts": r["status_counts"],
                        "controls": "ALL PASS" if r["controls"].all_pass else r["controls"].failed(),
                        "output_file_name": Path(r["path"]).name} for r in results]
            paths = [Path(r["path"]) for r in results]
        else:
            r = run_costco(remit, oi_path, out_dir)
            summary = [{"payment": r["payment"], "remittance_total": r["total"], "status_counts": r["status_counts"],
                        "nac": {region_: {"reference": ref, "amount": amt} for region_, (ref, amt) in r["nac"].items()},
                        "controls": "ALL PASS" if r["controls"].all_pass else r["controls"].failed(),
                        "output_file_name": Path(r["path"]).name}]
            paths = [Path(r["path"])]
    except HTTPException:
        shutil.rmtree(work, ignore_errors=True)
        raise
    except (ValueError, KeyError) as exc:          # bad / unexpected input file
        shutil.rmtree(work, ignore_errors=True)
        raise HTTPException(422, f"Could not process the file: {exc}") from exc
    except BaseException as exc:                   # includes SystemExit raised by the automation
        shutil.rmtree(work, ignore_errors=True)
        raise HTTPException(500, f"Processing failed: {exc}") from exc

    saved = []
    if LOCAL_OUTPUT is not None:
        LOCAL_OUTPUT.mkdir(exist_ok=True)
        for p in paths:
            shutil.copy2(p, LOCAL_OUTPUT / p.name)
            saved.append(f"Output/{p.name}")

    mode = (output or "").strip().lower()
    controls = "; ".join(s["controls"] if isinstance(s["controls"], str) else "FAILED: " + ", ".join(s["controls"])
                         for s in summary)
    info = {"status": "success", "customer": cust, "input_file": name, "controls": controls, "results": summary,
            "saved_locally": saved}
    if mode == "json":
        shutil.rmtree(work, ignore_errors=True)
        return JSONResponse(info)

    if len(paths) == 1:
        out_path, content_type = paths[0], XLSX
    else:                                                     # one Metro file holding several payments
        out_path, content_type = work / f"{Path(name).stem}_outputs.zip", "application/zip"
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for p in paths:
                z.write(p, p.name)
    if mode == "base64":            # JSON with the workbook inside (Oracle Visual Builder / other REST clients)
        file_b64 = base64.b64encode(out_path.read_bytes()).decode("ascii")
        shutil.rmtree(work, ignore_errors=True)
        return JSONResponse({**info, "output_file_name": out_path.name, "content_type": content_type,
                             "file_base64": file_b64})
    headers = {"X-Controls": controls[:500]}
    if saved:
        headers["X-Saved-To"] = "; ".join(saved)
    return FileResponse(out_path, media_type=content_type, filename=out_path.name, headers=headers,
                        background=BackgroundTask(shutil.rmtree, work, ignore_errors=True))
