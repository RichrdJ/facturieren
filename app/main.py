"""Facturieren: maandelijkse VM-export uploaden, per klant optellen en vergelijken met vorige maand."""
from __future__ import annotations

import base64
import io
import json
import os
import re
import secrets
import threading
from collections import Counter
from pathlib import Path

import openpyxl
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import excel

DATA_DIR = Path(os.environ.get("DATA_DIR", "/data"))
EXPORTS_DIR = DATA_DIR / "exports"
CUSTOMERS_FILE = DATA_DIR / "klanten.json"
STATIC_DIR = Path(__file__).parent / "static"
APP_USER = os.environ.get("APP_USER", "admin")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

MONTH_RE = re.compile(r"^\d{4}-\d{2}$")
LABELS = {
    "cpu": "CPU's",
    "memory": "Memory GB",
    "tier1_single": "tier1-single GB",
    "tier1_dual": "tier1-dual GB",
    "tier3_dual": "tier3-dual GB",
    "mgm": "mgm GB",
    "storage": "Total storage GB",
}

EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
app = FastAPI(title="Facturieren")
_lock = threading.Lock()
_cache: dict[str, tuple[float, excel.Export]] = {}


@app.middleware("http")
async def basic_auth(request: Request, call_next):
    if APP_PASSWORD and request.url.path != "/healthz":
        ok = False
        header = request.headers.get("authorization", "")
        if header.lower().startswith("basic "):
            try:
                user, _, pw = base64.b64decode(header[6:]).decode().partition(":")
                ok = secrets.compare_digest(user, APP_USER) and secrets.compare_digest(pw, APP_PASSWORD)
            except Exception:
                ok = False
        if not ok:
            return Response(status_code=401, headers={"WWW-Authenticate": 'Basic realm="Facturieren"'})
    return await call_next(request)


# ---------- opslag ----------

def _month_path(month: str) -> Path:
    if not MONTH_RE.match(month):
        raise HTTPException(400, "Maand moet de vorm JJJJ-MM hebben")
    return EXPORTS_DIR / f"{month}.xlsx"


def months() -> list[str]:
    return sorted(p.stem for p in EXPORTS_DIR.glob("*.xlsx") if MONTH_RE.match(p.stem))


def load(month: str) -> excel.Export:
    path = _month_path(month)
    if not path.exists():
        raise HTTPException(404, f"Geen export voor {month}")
    mtime = path.stat().st_mtime
    cached = _cache.get(month)
    if cached and cached[0] == mtime:
        return cached[1]
    export = excel.parse(path.read_bytes())
    _cache[month] = (mtime, export)
    return export


def previous(month: str) -> str | None:
    earlier = [m for m in months() if m < month]
    return earlier[-1] if earlier else None


def get_customers() -> list[dict]:
    if CUSTOMERS_FILE.exists():
        return json.loads(CUSTOMERS_FILE.read_text())
    return []


def month_from_filename(name: str) -> str | None:
    """'VMs export 01-10-2026.xlsx' -> '2026-10'."""
    m = re.search(r"(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})", name)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}"
    m = re.search(r"(\d{4})[-_.](\d{1,2})(?:[-_.]\d{1,2})?", name)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    return None


# ---------- berekeningen ----------

def totals(rows: list[dict]) -> dict:
    t = {k: sum(r[k] for r in rows) for k in excel.NUMERIC}
    t["vms"] = len(rows)
    return t


def select(export: excel.Export, pattern: str) -> list[dict]:
    return [r for r in export.rows if excel.matches(r["klant"], pattern)]


def diff(cur: list[dict], prev: list[dict]) -> dict:
    # Servernamen zijn niet altijd uniek; daarom per (naam, volgnummer) koppelen.
    def keyed(rows):
        seen: Counter = Counter()
        out = {}
        for r in rows:
            seen[r["servername"]] += 1
            out[(r["servername"], seen[r["servername"]])] = r
        return out

    c, p = keyed(cur), keyed(prev)
    added = [c[k] for k in c if k not in p]
    removed = [p[k] for k in p if k not in c]
    changed = []
    for k in c:
        if k in p:
            d = {f: c[k][f] - p[k][f] for f in excel.NUMERIC if c[k][f] != p[k][f]}
            moved = c[k]["klant"] != p[k]["klant"]
            if d or moved:
                changed.append({"servername": k[0], "klant": c[k]["klant"], "oud": p[k], "nieuw": c[k], "delta": d, "verplaatst": moved})
    return {"nieuw": added, "verwijderd": removed, "gewijzigd": changed}


def report(month: str, pattern: str) -> dict:
    export = load(month)
    rows = select(export, pattern)
    prev_month = previous(month)
    result = {"maand": month, "filter": pattern, "totaal": totals(rows), "vms": rows, "vorige_maand": prev_month}
    if prev_month:
        prev_rows = select(load(prev_month), pattern)
        result["vorig_totaal"] = totals(prev_rows)
        result["verschil"] = diff(rows, prev_rows)
    return result


def overview(month: str) -> dict:
    export = load(month)
    prev_month = previous(month)
    prev_export = load(prev_month) if prev_month else None
    regels = []
    covered = set()
    for c in get_customers():
        rows = select(export, c["filter"])
        covered.update(id(r) for r in rows)
        regel = {"naam": c["naam"], "filter": c["filter"], "totaal": totals(rows)}
        if prev_export:
            prev_rows = select(prev_export, c["filter"])
            regel["vorig_totaal"] = totals(prev_rows)
            d = diff(rows, prev_rows)
            regel["mutaties"] = {k: len(v) for k, v in d.items()}
        regels.append(regel)
    unassigned = [r for r in export.rows if id(r) not in covered]
    return {
        "maand": month,
        "vorige_maand": prev_month,
        "klanten": regels,
        "niet_toegewezen": totals(unassigned),
        "alles": totals(export.rows),
    }


# ---------- API ----------

@app.get("/healthz")
def healthz():
    return {"ok": True}


@app.get("/api/months")
def api_months():
    out = []
    for m in months():
        try:
            e = load(m)
            out.append({"maand": m, "werkblad": e.sheet, "vms": len(e.rows), "waarschuwingen": e.warnings})
        except excel.ExportError as err:
            out.append({"maand": m, "fout": str(err)})
    return out


@app.post("/api/upload")
async def api_upload(file: UploadFile = File(...), maand: str = Form("")):
    data = await file.read()
    month = maand.strip() or month_from_filename(file.filename or "")
    if not month:
        raise HTTPException(400, "Kon de maand niet uit de bestandsnaam halen; kies de maand handmatig.")
    path = _month_path(month)
    try:
        export = excel.parse(data, file.filename or "")
        if not excel.is_xlsx(data):
            # Kale CSV: bewaar als nette xlsx.
            data = excel.clean_xlsx(data, file.filename or "")
    except excel.ExportError as err:
        raise HTTPException(400, str(err))
    with _lock:
        existed = path.exists()
        path.write_bytes(data)
        _cache.pop(month, None)
    return {"maand": month, "vms": len(export.rows), "overschreven": existed, "waarschuwingen": export.warnings}


@app.delete("/api/months/{month}")
def api_delete_month(month: str):
    path = _month_path(month)
    if not path.exists():
        raise HTTPException(404, "Niet gevonden")
    path.unlink()
    _cache.pop(month, None)
    return {"ok": True}


@app.get("/api/months/{month}/xlsx")
def api_month_xlsx(month: str):
    """De export van die maand als nette xlsx (getallen als getal, opgemaakte kop, filter)."""
    path = _month_path(month)
    if not path.exists():
        raise HTTPException(404, "Niet gevonden")
    y, m = month.split("-")
    return Response(
        excel.clean_xlsx(path.read_bytes()),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="VMs export 01-{m}-{y}.xlsx"'},
    )


@app.post("/api/convert")
async def api_convert(file: UploadFile = File(...)):
    """Losse omzetting csv/xlsx -> nette xlsx, zonder iets op te slaan."""
    data = await file.read()
    try:
        out = excel.clean_xlsx(data, file.filename or "")
    except excel.ExportError as err:
        raise HTTPException(400, str(err))
    month = month_from_filename(file.filename or "")
    name = f"VMs export 01-{month[5:]}-{month[:4]}.xlsx" if month else re.sub(r"\.[^.]*$", "", file.filename or "export") + ".xlsx"
    return Response(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{name}"'},
    )


@app.get("/api/folders")
def api_folders(maand: str):
    """Alle klantmappen (zonder de datacenter-root) met aantal VM's, voor suggesties."""
    count: Counter = Counter()
    for r in load(maand).rows:
        segs = [s for s in r["klant"].split("/") if s][1:]
        for i in range(1, len(segs) + 1):
            count["/".join(segs[:i])] += 1
    return [{"map": k, "vms": v} for k, v in sorted(count.items())]


class Customer(BaseModel):
    naam: str
    filter: str


@app.get("/api/customers")
def api_customers():
    return get_customers()


@app.put("/api/customers")
def api_save_customers(customers: list[Customer]):
    data = [c.model_dump() for c in customers if c.naam.strip() and c.filter.strip()]
    with _lock:
        CUSTOMERS_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False))
    return data


@app.get("/api/report")
def api_report(maand: str, filter: str):
    return report(maand, filter)


@app.get("/api/overview")
def api_overview(maand: str):
    return overview(maand)


@app.get("/api/export")
def api_export(maand: str):
    """Excel met overzicht per klant + per klant een tabblad met de VM's."""
    ov = overview(maand)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overzicht"
    keys = ["vms", *excel.NUMERIC]
    head = ["Klant", "Filter", "VM's", *LABELS.values()]
    if ov["vorige_maand"]:
        head += [f"Δ {h}" for h in ["VM's", *LABELS.values()]]
    ws.append(head)
    for k in ov["klanten"]:
        row = [k["naam"], k["filter"], *[k["totaal"][x] for x in keys]]
        if "vorig_totaal" in k:
            row += [k["totaal"][x] - k["vorig_totaal"][x] for x in keys]
        ws.append(row)
    ws.append(["Niet toegewezen", "", *[ov["niet_toegewezen"][x] for x in keys]])
    used = set()
    for k in ov["klanten"]:
        title = re.sub(r"[\[\]:*?/\\]", "_", k["naam"])[:31]
        while title in used:
            title = title[:28] + str(len(used))
        used.add(title)
        s = wb.create_sheet(title)
        s.append(["servername", "klant", "cluster", *LABELS.values()])
        for r in select(load(maand), k["filter"]):
            s.append([r["servername"], r["klant"], r["cluster"], *[r[x] for x in excel.NUMERIC]])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="facturatie {maand}.xlsx"'},
    )


@app.get("/")
def index():
    return FileResponse(STATIC_DIR / "index.html")


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
