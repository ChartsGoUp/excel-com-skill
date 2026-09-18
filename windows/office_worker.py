from __future__ import annotations

import argparse
import json
from pathlib import Path


def emit(data: dict, code: int = 0) -> int:
    print(json.dumps(data, indent=2, sort_keys=True, default=str))
    return code


def require_pywin32():
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
        return pythoncom, win32com.client
    except Exception as exc:
        raise RuntimeError(
            "pywin32 is required in the Windows Python environment. "
            "Run windows/bootstrap.ps1."
        ) from exc


def office_health() -> dict:
    pythoncom, win32 = require_pywin32()
    result: dict = {"status": "PASS", "word": {}, "excel": {}}
    pythoncom.CoInitialize()
    try:
        word = None
        excel = None
        try:
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            info: dict = {"available": True, "version": str(word.Version)}
            # Instantiate-only success is not enough: an unlicensed Office
            # (grace expired) still starts but cannot open/save files. Probe
            # a real Add/Save/Open round-trip so engine-status never reports
            # a false PASS.
            try:
                _probe_word_file_ops(word)
                info["file_operations"] = True
            except Exception as exc:
                info["file_operations"] = False
                info["file_error"] = str(exc)
                result["status"] = "FAIL"
            result["word"] = info
        except Exception as exc:
            result["word"] = {"available": False, "error": str(exc)}
            result["status"] = "FAIL"
        finally:
            if word is not None:
                try:
                    word.Quit(False)
                except Exception:
                    pass

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            info = {"available": True, "version": str(excel.Version)}
            try:
                _probe_excel_file_ops(excel)
                info["file_operations"] = True
            except Exception as exc:
                info["file_operations"] = False
                info["file_error"] = str(exc)
                result["status"] = "FAIL"
            result["excel"] = info
        except Exception as exc:
            result["excel"] = {"available": False, "error": str(exc)}
            result["status"] = "FAIL"
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
    finally:
        pythoncom.CoUninitialize()
    return result


def _probe_word_file_ops(word) -> None:
    import tempfile

    tmpdir = Path(tempfile.gettempdir())
    probe = tmpdir / "office_health_probe.docx"
    doc = None
    try:
        if probe.exists():
            probe.unlink()
    except Exception:
        pass
    try:
        doc = word.Documents.Add()
        if doc is None:
            raise RuntimeError("Word Documents.Add returned None.")
        doc.Content.Text = "Office health probe."
        doc.SaveAs(str(probe))
        doc.Close(False)
        doc = None
        reopened = word.Documents.Open(str(probe), ReadOnly=True, AddToRecentFiles=False)
        _require_opened(reopened, probe, "Word")
        reopened.Close(False)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            if probe.exists():
                probe.unlink()
        except Exception:
            pass


def _probe_excel_file_ops(excel) -> None:
    import tempfile

    tmpdir = Path(tempfile.gettempdir())
    probe = tmpdir / "office_health_probe.xlsx"
    wb = None
    try:
        if probe.exists():
            probe.unlink()
    except Exception:
        pass
    try:
        wb = excel.Workbooks.Add()
        if wb is None:
            raise RuntimeError("Excel Workbooks.Add returned None.")
        wb.ActiveSheet.Range("A1").Value = "Office health probe."
        wb.SaveAs(str(probe))
        wb.Close(False)
        wb = None
        reopened = excel.Workbooks.Open(str(probe), UpdateLinks=0, ReadOnly=True)
        _require_opened(reopened, probe, "Excel")
        reopened.Close(False)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        try:
            if probe.exists():
                probe.unlink()
        except Exception:
            pass


def _require_opened(doc, src: Path, app: str) -> None:
    if doc is not None:
        return
    raise RuntimeError(
        f"{app} COM opened no document for: {src}. "
        "Word/Excel started but Documents.Open/Workbooks.Open returned None. "
        "Common causes: unlicensed Office (grace expired, reduced-functionality mode), "
        "a blocking activation dialog, Trust Center blocking the file location, "
        "or an unreadable path. Check licensing with "
        "cscript \"C:\\Program Files\\Microsoft Office\\Root\\Office16\\OSPP.VBS\" /dstatus."
    )


def _word_pdf(src: Path, pdf: Path) -> dict:
    pythoncom, win32 = require_pywin32()
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
        _require_opened(doc, src, "Word")
        doc.Repaginate()
        doc.ExportAsFixedFormat(str(pdf), 17)
        return {"status": "PASS", "engine": "word-com", "pdf": str(pdf)}
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _excel_pdf(src: Path, pdf: Path) -> dict:
    pythoncom, win32 = require_pywin32()
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        wb = excel.Workbooks.Open(
            str(src),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        _require_opened(wb, src, "Excel")
        wb.ExportAsFixedFormat(0, str(pdf))
        return {"status": "PASS", "engine": "excel-com", "pdf": str(pdf)}
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def render(src: Path, out: Path) -> dict:
    src = src.resolve()
    out = out.resolve()
    out.mkdir(parents=True, exist_ok=True)
    pdf = out / f"{src.stem}.pdf"
    suffix = src.suffix.lower()
    if suffix in {".docx", ".docm", ".doc"}:
        return _word_pdf(src, pdf)
    if suffix in {".xlsx", ".xlsm", ".xlsb", ".xls"}:
        return _excel_pdf(src, pdf)
    return {"status": "FAIL", "message": f"Unsupported Office file: {src}"}


def recalculate_xlsx(src: Path) -> dict:
    pythoncom, win32 = require_pywin32()
    src = src.resolve()
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        wb = excel.Workbooks.Open(
            str(src), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True
        )
        _require_opened(wb, src, "Excel")
        excel.CalculateFullRebuild()
        wb.Save()
        return {"status": "PASS", "engine": "excel-com", "file": str(src)}
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def autofit_xlsx(src: Path) -> dict:
    pythoncom, win32 = require_pywin32()
    src = src.resolve()
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        wb = excel.Workbooks.Open(
            str(src), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True
        )
        _require_opened(wb, src, "Excel")
        changed = []
        for ws in wb.Worksheets:
            used = ws.UsedRange
            used.Columns.AutoFit()
            used.Rows.AutoFit()
            changed.append(ws.Name)
        wb.Save()
        return {
            "status": "PASS",
            "engine": "excel-com",
            "file": str(src),
            "sheets": changed,
        }
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def repaginate_docx(src: Path) -> dict:
    pythoncom, win32 = require_pywin32()
    src = src.resolve()
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(src), ReadOnly=False, AddToRecentFiles=False)
        _require_opened(doc, src, "Word")
        try:
            doc.Fields.Update()
        except Exception:
            pass
        try:
            for toc in doc.TablesOfContents:
                toc.Update()
        except Exception:
            pass
        doc.Repaginate()
        doc.Save()
        return {"status": "PASS", "engine": "word-com", "file": str(src)}
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def snapshot_xlsx(src: Path) -> dict:
    pythoncom, win32 = require_pywin32()
    src = src.resolve()
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        wb = excel.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True)
        _require_opened(wb, src, "Excel")
        sheets = []
        for ws in wb.Worksheets:
            used = ws.UsedRange
            sheets.append(
                {
                    "name": ws.Name,
                    "visible": int(ws.Visible),
                    "used_range": used.Address,
                    "rows": int(used.Rows.Count),
                    "columns": int(used.Columns.Count),
                    "page_orientation": int(ws.PageSetup.Orientation),
                    "zoom": ws.Application.ActiveWindow.Zoom if ws.Application.ActiveWindow else None,
                }
            )
        return {
            "status": "PASS",
            "engine": "excel-com",
            "file": str(src),
            "sheets": sheets,
        }
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def snapshot_docx(src: Path) -> dict:
    pythoncom, win32 = require_pywin32()
    src = src.resolve()
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
        _require_opened(doc, src, "Word")
        doc.Repaginate()
        sections = []
        for sec in doc.Sections:
            setup = sec.PageSetup
            sections.append(
                {
                    "orientation": int(setup.Orientation),
                    "page_width": float(setup.PageWidth),
                    "page_height": float(setup.PageHeight),
                    "top_margin": float(setup.TopMargin),
                    "bottom_margin": float(setup.BottomMargin),
                    "left_margin": float(setup.LeftMargin),
                    "right_margin": float(setup.RightMargin),
                }
            )
        return {
            "status": "PASS",
            "engine": "word-com",
            "file": str(src),
            "pages": int(doc.ComputeStatistics(2)),
            "paragraphs": int(doc.Paragraphs.Count),
            "tables": int(doc.Tables.Count),
            "sections": sections,
        }
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Native Microsoft Office COM worker")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("health")

    p = sub.add_parser("render")
    p.add_argument("--input", required=True)
    p.add_argument("--out", required=True)

    for name in ("recalculate-xlsx", "autofit-xlsx", "repaginate-docx", "snapshot"):
        p = sub.add_parser(name)
        p.add_argument("--input", required=True)

    args = parser.parse_args(argv)
    try:
        if args.cmd == "health":
            data = office_health()
        elif args.cmd == "render":
            data = render(Path(args.input), Path(args.out))
        elif args.cmd == "recalculate-xlsx":
            data = recalculate_xlsx(Path(args.input))
        elif args.cmd == "autofit-xlsx":
            data = autofit_xlsx(Path(args.input))
        elif args.cmd == "repaginate-docx":
            data = repaginate_docx(Path(args.input))
        elif args.cmd == "snapshot":
            src = Path(args.input)
            if src.suffix.lower() in {".xlsx", ".xlsm", ".xlsb", ".xls"}:
                data = snapshot_xlsx(src)
            else:
                data = snapshot_docx(src)
        else:
            raise AssertionError(args.cmd)
    except Exception as exc:
        data = {"status": "FAIL", "error": type(exc).__name__, "message": str(exc)}
    return emit(data, 0 if data.get("status") == "PASS" else 2)


if __name__ == "__main__":
    # Stable entrypoint used by bootstrap/transport; execute hardened implementation.
    # When imported as `office_worker`, the legacy helpers above remain available.
    from office_worker_hardened import main as hardened_main
    raise SystemExit(hardened_main())
