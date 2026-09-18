from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_ops import compare_docx, validate_docx
from .xlsx_ops import compare_xlsx, validate_xlsx
from .render import render_office


def _overall(validation: dict[str, Any], comparison: dict[str, Any] | None, render: dict[str, Any] | None) -> str:
    if validation.get("status") == "FAIL":
        return "FAIL"
    if comparison is not None and not comparison.get("content_equal", False):
        return "FAIL"
    if render is not None and render.get("status") == "FAIL":
        return "FAIL"
    if validation.get("status") == "WARN" or (render is not None and render.get("status") == "UNAVAILABLE"):
        return "WARN"
    return "PASS"


def qa_docx(source: str | Path, candidate: str | Path, out_dir: str | Path | None = None) -> dict[str, Any]:
    validation = validate_docx(candidate)
    comparison = compare_docx(source, candidate)
    render = render_office(candidate, out_dir) if out_dir else None
    return {
        "artifact_type": "docx",
        "source": str(source),
        "candidate": str(candidate),
        "validation": validation,
        "comparison": comparison,
        "render": render,
        "visual_inspection_required": True,
        "status": _overall(validation, comparison, render),
    }


def qa_xlsx(source: str | Path, candidate: str | Path, out_dir: str | Path | None = None) -> dict[str, Any]:
    validation = validate_xlsx(candidate)
    comparison = compare_xlsx(source, candidate)
    render = render_office(candidate, out_dir) if out_dir else None
    return {
        "artifact_type": "xlsx",
        "source": str(source),
        "candidate": str(candidate),
        "validation": validation,
        "comparison": comparison,
        "render": render,
        "visual_inspection_required": True,
        "status": _overall(validation, comparison, render),
    }
