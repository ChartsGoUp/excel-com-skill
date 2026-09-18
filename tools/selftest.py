#!/usr/bin/env python3
from pathlib import Path
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font

from office_artifact_system.docx_ops import validate_docx, compare_docx
from office_artifact_system.xlsx_ops import validate_xlsx, compare_xlsx, inspect_xlsx
from office_artifact_system.xlsx_style import financial_model_style, semantic_style


def main() -> int:
    style = financial_model_style()
    assert style["font"]["name"] == "Arial"
    assert semantic_style("hardcode_input")["font_color"] == "0000FF"

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        d1 = td / "a.docx"
        d2 = td / "b.docx"
        doc = Document()
        p = doc.add_paragraph("Test body")
        p.runs[0].font.name = "Arial"
        p.runs[0].font.size = Pt(10)
        doc.save(d1)
        doc.save(d2)
        assert validate_docx(d1)["errors"] == 0
        assert compare_docx(d1, d2)["content_equal"]

        x1 = td / "a.xlsx"
        x2 = td / "b.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Model"
        ws["A1"] = "Label"
        ws["A1"].font = Font(name="Arial", size=10)
        ws["B1"] = 1
        ws["B1"].font = Font(name="Arial", size=10)
        ws["B2"] = "=B1+1"
        ws["B2"].font = Font(name="Arial", size=10)
        ws.freeze_panes = "B2"
        wb.save(x1)
        wb.save(x2)

        inspection = inspect_xlsx(x1)
        assert inspection["edit_preflight"]["risk"] == "LOW"
        assert validate_xlsx(x1)["errors"] == 0
        assert compare_xlsx(x1, x2)["content_equal"]

        # Comparison must catch formatting/structure drift, not only cell values.
        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = "Model"
        ws2["A1"] = "Label"
        ws2["A1"].font = Font(name="Calibri", size=11)
        ws2["B1"] = 1
        ws2["B1"].font = Font(name="Arial", size=10)
        ws2["B2"] = "=B1+1"
        ws2["B2"].font = Font(name="Arial", size=10)
        ws2.freeze_panes = "C3"
        wb2.save(x2)

        comparison = compare_xlsx(x1, x2)
        assert not comparison["content_equal"]
        types = {d["type"] for d in comparison["differences"]}
        assert "style" in types
        assert "sheet_structure" in types

    print("selftest: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
