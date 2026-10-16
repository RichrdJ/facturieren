"""VM-overzicht voor een klant op basis van RVTools-exports (los van de chargeback-facturatie).

Zelfde logica als het PowerShell-script "<Klant> VM export v1.1.ps1":
- per RVTools-export de tabbladen vInfo, vCPU en vDisk;
- Folder komt uit vCPU, gekoppeld op VM-naam;
- een VM hoort bij de klant als Folder of het attribuut HIAB_BU_Separation de klantnaam bevat
  (sommige platforms vullen Folder niet; dan is HIAB_BU_Separation de enige aanwijzing);
- Memory = MiB / 1024 afgerond, Storage = som van "Capacity MiB" uit vDisk / 1024, naar boven afgerond;
- dubbele VM's (over exports heen) worden samengevoegd.
"""
from __future__ import annotations

import io
import math
import re

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .excel import ExportError, _open

COLUMNS = ["VM", "Klant", "CPUs", "Memory", "Storage", "IPv4 Address"]


def is_rvtools(data: bytes, name: str = "") -> bool:
    """Herkent een RVTools-export aan de naam of aan het tabblad vInfo."""
    if re.match(r"rvtools", name or "", re.I):
        return True
    if data[:2] != b"PK":
        return False
    try:
        wb = _open(data)
        names = set(wb.sheetnames)
        wb.close()
        return {"vInfo", "vCPU", "vDisk"} <= names
    except ExportError:
        return False


def _sheet(wb, name: str) -> list[dict]:
    if name not in wb.sheetnames:
        raise ExportError(f"Tabblad '{name}' ontbreekt; is dit wel een RVTools-export?")
    it = wb[name].iter_rows(values_only=True)
    try:
        header = [str(h).strip() if h is not None else "" for h in next(it)]
    except StopIteration:
        return []
    rows = []
    for raw in it:
        row = {}
        for h, v in zip(header, raw):
            if h and h not in row:  # bij dubbele kolomnamen telt de eerste
                row[h] = v
        rows.append(row)
    return rows


def _num(v) -> float:
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def parse(files: list[tuple[str, bytes]], klant: str) -> dict:
    """Geeft {"vms": [...], "bestanden": [...], "waarschuwingen": [...]} voor één klant."""
    pattern = re.compile(re.escape(klant.strip()), re.I)
    found: dict[str, dict] = {}
    bestanden, warnings = [], []
    for name, data in files:
        try:
            wb = _open(data)
            v_info, v_cpu, v_disk = (_sheet(wb, s) for s in ("vInfo", "vCPU", "vDisk"))
            wb.close()
        except ExportError as err:
            warnings.append(f"{name}: {err}")
            continue

        # Opzoeken op VM-naam, hoofdletterongevoelig (zoals PowerShell-hashtables).
        folders: dict[str, str] = {}
        for r in v_cpu:
            vm = _text(r.get("VM"))
            if vm and vm.lower() not in folders:
                folders[vm.lower()] = _text(r.get("Folder"))
        storage: dict[str, float] = {}
        for r in v_disk:
            vm = _text(r.get("VM"))
            if vm:
                storage[vm.lower()] = storage.get(vm.lower(), 0.0) + _num(r.get("Capacity MiB"))

        count = 0
        for r in v_info:
            vm = _text(r.get("VM"))
            if not vm:
                continue
            folder = folders.get(vm.lower(), "")
            bu = _text(r.get("HIAB_BU_Separation"))
            if not (pattern.search(folder) or pattern.search(bu)):
                continue
            mem = _num(r.get("Memory"))
            mib = storage.get(vm.lower(), 0.0)
            row = {
                "VM": vm,
                "Klant": folder or bu,
                "CPUs": int(_num(r.get("CPUs"))),
                "Memory": int(round(mem / 1024)) if mem else 0,
                "Storage": int(math.ceil(mib / 1024)) if mib else 0,
                "IPv4 Address": _text(r.get("Primary IP Address")),
                "Powerstate": _text(r.get("Powerstate")),
                "Bron": name,
            }
            found.setdefault(vm.lower(), row)  # dubbele VM's: de eerste telt
            count += 1
        bestanden.append({"naam": name, "vms": count})

    vms = sorted(found.values(), key=lambda r: r["VM"].lower())
    return {"vms": vms, "bestanden": bestanden, "waarschuwingen": warnings}


def to_xlsx(vms: list[dict]) -> bytes:
    """Output in de opmaak van het klantbestand: Verdana 9, kopregel vet wit op zwart."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Blad1"
    ws.append(COLUMNS)
    for r in vms:
        ws.append([r.get("VM", ""), r.get("Klant", ""), r.get("CPUs", 0), r.get("Memory", 0),
                   r.get("Storage", 0), str(r.get("IPv4 Address", "") or "")])
    head_font = Font(name="Verdana", size=9, bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="000000")
    for cell in ws[1]:
        cell.font = head_font
        cell.fill = fill
        cell.alignment = Alignment(horizontal="left")
    body_font = Font(name="Verdana", size=9)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
        row[5].number_format = "@"  # IP als tekst, zodat Excel er geen getal van maakt
    for i, col in enumerate(ws.iter_cols(values_only=True), start=1):
        width = max((len(str(v)) for v in col if v is not None), default=6) * 1.1 + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 70)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
