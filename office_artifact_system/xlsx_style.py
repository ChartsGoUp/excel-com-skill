from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path

from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


PROJECT_ROOT = Path(__file__).resolve().parents[1]
STYLE_CONFIG = PROJECT_ROOT / "config" / "financial-model-style.json"

NAVY = "1F4E79"
LIGHT_BLUE = "D9EAF7"
LIGHT_GREY = "E7E6E6"
INPUT_YELLOW = "FFF2CC"
WHITE = "FFFFFF"
BLACK = "000000"
BLUE_INPUT = "0000FF"


@lru_cache(maxsize=1)
def financial_model_style() -> dict:
    return json.loads(STYLE_CONFIG.read_text(encoding="utf-8"))


def semantic_style(name: str) -> dict:
    styles = financial_model_style().get("semantic_styles", {})
    if name not in styles:
        raise KeyError(f"Unknown financial-model semantic style: {name}")
    return dict(styles[name])


def apply_default_font(ws, name: str | None = None, size: int | None = None) -> None:
    cfg = financial_model_style()["font"]
    name = name or cfg["name"]
    size = size or cfg["body_pt"]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            current = cell.font
            cell.font = Font(
                name=name,
                size=size,
                bold=current.bold,
                italic=current.italic,
                color=current.color,
                underline=current.underline,
            )


def _side(spec: str | None):
    if not spec:
        return Side(style=None)
    return Side(style=spec, color="000000")


def apply_semantic_style(cell, name: str) -> None:
    spec = semantic_style(name)
    font_cfg = financial_model_style()["font"]
    cell.font = Font(
        name=font_cfg["name"],
        size=font_cfg["body_pt"],
        bold=bool(spec.get("bold", False)),
        italic=bool(spec.get("italic", False)),
        color=spec.get("font_color", BLACK),
    )
    if spec.get("fill"):
        cell.fill = PatternFill("solid", fgColor=spec["fill"])
    if spec.get("align"):
        cell.alignment = Alignment(horizontal=spec["align"], vertical="center")
    if spec.get("top_border") or spec.get("bottom_border"):
        cell.border = Border(
            top=_side(spec.get("top_border")),
            bottom=_side(spec.get("bottom_border")),
        )


def style_section_bar(cells, fill: str | None = None) -> None:
    fill = fill or semantic_style("section_header")["fill"]
    for cell in cells:
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="left", vertical="center")


def style_input(cell) -> None:
    apply_semantic_style(cell, "hardcode_input")


def financial_number_format(decimals: int = 0) -> str:
    formats = financial_model_style()["number_formats"]
    return formats["integer"] if decimals <= 0 else formats["decimal_1"]


def percentage_format(decimals: int = 1) -> str:
    formats = financial_model_style()["number_formats"]
    return formats["percentage_1"] if decimals == 1 else f'0.{"0"*decimals}%;(0.{"0"*decimals}%);-'


def set_view(ws, zoom: int | None = None, show_gridlines: bool | None = None, freeze: str | None = None) -> None:
    views = financial_model_style()["views"]
    ws.sheet_view.zoomScale = zoom if zoom is not None else views["standard_zoom"]
    ws.sheet_view.showGridLines = show_gridlines if show_gridlines is not None else views["gridlines"]
    if freeze:
        ws.freeze_panes = freeze


def fit_columns(ws, min_width: float = 8, max_width: float | None = None, padding: float = 2) -> None:
    max_width = max_width or financial_model_style()["columns"]["autofit_max_expansion_width"]
    for idx in range(1, ws.max_column + 1):
        letter = get_column_letter(idx)
        dim = ws.column_dimensions[letter]
        if dim.hidden or (dim.width is not None and dim.width <= 3):
            continue
        max_len = 0
        for cell in ws[letter]:
            if cell.value is None:
                continue
            max_len = max(max_len, max(len(part) for part in str(cell.value).splitlines()))
        estimated = max(min_width, min(max_width, max_len + padding))
        # Expand only. Native Excel COM performs the authoritative final fit.
        if dim.width is None or estimated > dim.width:
            dim.width = estimated
