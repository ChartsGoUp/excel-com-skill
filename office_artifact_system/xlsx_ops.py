from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any
import re

from openpyxl import load_workbook

from .common import Finding, hash_file

ERROR_VALUES = {"#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}


def _cell_payload(ws):
    out = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                out[c.coordinate] = c.value
    return out


def inspect_xlsx(path: str | Path) -> dict[str, Any]:
    path = Path(path)
    wb = load_workbook(path, data_only=False, read_only=False)
    sheets = []
    for ws in wb.worksheets:
        formula_count = 0
        populated = 0
        errors = []
        fonts = Counter()
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                populated += 1
                if c.data_type == "f" or (isinstance(c.value, str) and c.value.startswith("=")):
                    formula_count += 1
                if c.data_type == "e" or (isinstance(c.value, str) and c.value in ERROR_VALUES):
                    errors.append(c.coordinate)
                if c.font and c.font.name:
                    fonts[c.font.name] += 1
        sheets.append({
            "title": ws.title,
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "populated_cells": populated,
            "formula_cells": formula_count,
            "error_cells": errors,
            "show_gridlines": bool(ws.sheet_view.showGridLines),
            "freeze_panes": str(ws.freeze_panes) if ws.freeze_panes else None,
            "zoom": ws.sheet_view.zoomScale,
            "sheet_state": ws.sheet_state,
            "tab_color": getattr(ws.sheet_properties.tabColor, "rgb", None) if ws.sheet_properties.tabColor else None,
            "fonts": dict(fonts.most_common()),
        })
    links = []
    for link in getattr(wb, "_external_links", []) or []:
        links.append(str(getattr(link, "file_link", None) or link))
    return {
        "path": str(path),
        "sha256": hash_file(path),
        "sheet_names": wb.sheetnames,
        "defined_names": sorted(str(n) for n in wb.defined_names),
        "external_link_count": len(links),
        "external_links": links,
        "sheets": sheets,
    }


def validate_xlsx(path: str | Path, required_font: str = "Arial") -> dict[str, Any]:
    path = Path(path)
    wb = load_workbook(path, data_only=False, read_only=False)
    findings: list[Finding] = []
    if not wb.sheetnames:
        findings.append(Finding("error", "XLSX_NO_SHEETS", "Workbook contains no worksheets."))

    if getattr(wb, "_external_links", None):
        findings.append(Finding("warning", "XLSX_EXTERNAL_LINK", f"Workbook contains {len(wb._external_links)} external link record(s)."))

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if c.data_type == "e" or (isinstance(c.value, str) and c.value in ERROR_VALUES):
                    findings.append(Finding("error", "XLSX_ERROR_CELL", f"Spreadsheet error value {c.value!r}.", f"{ws.title}!{c.coordinate}"))
                if c.font and c.font.name and c.font.name.lower() != required_font.lower():
                    findings.append(Finding("warning", "XLSX_FONT", f"Font {c.font.name!r}; expected {required_font}.", f"{ws.title}!{c.coordinate}"))
                if isinstance(c.value, str) and c.value.startswith("=") and "#REF!" in c.value.upper():
                    findings.append(Finding("error", "XLSX_BROKEN_FORMULA", "Formula contains #REF!.", f"{ws.title}!{c.coordinate}"))

        if ws.sheet_state == "visible" and ws.max_row > 5 and ws.sheet_view.zoomScale and not (70 <= ws.sheet_view.zoomScale <= 130):
            findings.append(Finding("warning", "XLSX_ZOOM", f"Unusual worksheet zoom {ws.sheet_view.zoomScale}%.", ws.title))
        for key, dim in ws.column_dimensions.items():
            if dim.width and dim.width > 60:
                findings.append(Finding("warning", "XLSX_WIDE_COLUMN", f"Column width {dim.width} is unusually wide.", f"{ws.title}!{key}:{key}"))

    if len(findings) > 500:
        counts = Counter((f.severity, f.code, f.message) for f in findings)
        compact = []
        seen = set()
        for f in findings:
            key = (f.severity, f.code, f.message)
            if key in seen:
                continue
            seen.add(key)
            n = counts[key]
            compact.append(Finding(f.severity, f.code, f"{f.message} Occurrences: {n}.", f.location))
        findings = compact

    errors = sum(f.severity == "error" for f in findings)
    warnings = sum(f.severity == "warning" for f in findings)
    return {
        "path": str(path),
        "sha256": hash_file(path),
        "status": "FAIL" if errors else ("WARN" if warnings else "PASS"),
        "errors": errors,
        "warnings": warnings,
        "findings": [f.to_dict() for f in findings],
    }


def compare_xlsx(source: str | Path, candidate: str | Path) -> dict[str, Any]:
    swb = load_workbook(source, data_only=False, read_only=False)
    cwb = load_workbook(candidate, data_only=False, read_only=False)
    diffs = []
    if swb.sheetnames != cwb.sheetnames:
        diffs.append({"type": "sheet_names", "source": swb.sheetnames, "candidate": cwb.sheetnames})
    common = [s for s in swb.sheetnames if s in cwb.sheetnames]
    for name in common:
        sp = _cell_payload(swb[name])
        cp = _cell_payload(cwb[name])
        coords = sorted(set(sp) | set(cp))
        for coord in coords:
            if sp.get(coord) != cp.get(coord):
                diffs.append({"type": "cell", "sheet": name, "cell": coord, "source": sp.get(coord), "candidate": cp.get(coord)})
                if len(diffs) >= 100:
                    break
        if len(diffs) >= 100:
            break
    return {
        "source": str(source),
        "candidate": str(candidate),
        "source_sha256": hash_file(source),
        "candidate_sha256": hash_file(candidate),
        "content_equal": not diffs,
        "differences": diffs,
    }
