from __future__ import annotations

import argparse
import ctypes
import gc
import json
import os
from contextlib import contextmanager
from pathlib import Path
import tempfile
import time

import office_worker as legacy

XL_DONE = 0
FORCE_DISABLE_MACROS = 3
WAIT_OBJECT_0 = 0
WAIT_TIMEOUT = 258
SYNCHRONIZE = 0x00100000
PROCESS_TERMINATE = 0x0001
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


class FILETIME(ctypes.Structure):
    _fields_ = [("dwLowDateTime", ctypes.c_uint32), ("dwHighDateTime", ctypes.c_uint32)]


def _creation_filetime(handle) -> int | None:
    if not handle:
        return None
    creation = FILETIME(); exit_time = FILETIME(); kernel = FILETIME(); user = FILETIME()
    ok = ctypes.windll.kernel32.GetProcessTimes(
        handle, ctypes.byref(creation), ctypes.byref(exit_time), ctypes.byref(kernel), ctypes.byref(user)
    )
    if not ok:
        return None
    return (int(creation.dwHighDateTime) << 32) | int(creation.dwLowDateTime)


def _ownership_path() -> Path | None:
    value = os.environ.get("OFFICE_COM_OWNERSHIP_FILE")
    return Path(value) if value else None


def _publish_ownership(app_name: str, pid: int, handle) -> None:
    target = _ownership_path()
    if not target:
        return
    payload = {
        "app": app_name,
        "pid": int(pid),
        "creation_filetime": _creation_filetime(handle),
        "worker_pid": os.getpid(),
    }
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload), encoding="utf-8")
    os.replace(tmp, target)


def _clear_ownership() -> None:
    target = _ownership_path()
    if target:
        try:
            target.unlink(missing_ok=True)
        except Exception:
            pass


def _open_process(pid: int):
    return ctypes.windll.kernel32.OpenProcess(
        SYNCHRONIZE | PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid)
    ) or None


def _wait_or_terminate(handle, pid: int, grace_ms: int = 4000) -> dict:
    meta = {"owned_pid": int(pid), "forced_termination": False}
    if not handle:
        meta["cleanup_warning"] = "Could not open an ownership handle for the Office process."
        return meta
    kernel32 = ctypes.windll.kernel32
    try:
        state = kernel32.WaitForSingleObject(handle, grace_ms)
        if state == WAIT_TIMEOUT:
            if kernel32.TerminateProcess(handle, 1):
                meta["forced_termination"] = True
                kernel32.WaitForSingleObject(handle, 5000)
            else:
                meta["cleanup_warning"] = f"Owned Office PID {pid} did not exit and could not be terminated."
        elif state != WAIT_OBJECT_0:
            meta["cleanup_warning"] = f"Unexpected wait result {state} for owned Office PID {pid}."
    finally:
        kernel32.CloseHandle(handle)
    return meta


@contextmanager
def excel_session():
    pythoncom, win32 = legacy.require_pywin32()
    import win32process  # type: ignore
    pythoncom.CoInitialize()
    app = None; handle = None; meta: dict = {}
    try:
        app = win32.DispatchEx("Excel.Application")
        pid = int(win32process.GetWindowThreadProcessId(app.Hwnd)[1])
        handle = _open_process(pid)
        meta["owned_pid"] = pid
        _publish_ownership("excel", pid, handle)
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.ScreenUpdating = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = FORCE_DISABLE_MACROS
        yield app, meta
    finally:
        if app is not None:
            try:
                for idx in range(app.Workbooks.Count, 0, -1):
                    try: app.Workbooks(idx).Close(False)
                    except Exception: pass
            except Exception:
                pass
            try: app.Quit()
            except Exception: pass
        app = None
        gc.collect()
        if meta.get("owned_pid"):
            meta.update(_wait_or_terminate(handle, int(meta["owned_pid"])))
        elif handle:
            ctypes.windll.kernel32.CloseHandle(handle)
        _clear_ownership()
        pythoncom.CoUninitialize()


@contextmanager
def word_session():
    pythoncom, win32 = legacy.require_pywin32()
    import win32process  # type: ignore
    pythoncom.CoInitialize()
    app = None; handle = None; meta: dict = {}
    try:
        app = win32.DispatchEx("Word.Application")
        pid = int(win32process.GetWindowThreadProcessId(app.Hwnd)[1])
        handle = _open_process(pid)
        meta["owned_pid"] = pid
        _publish_ownership("word", pid, handle)
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = FORCE_DISABLE_MACROS
        yield app, meta
    finally:
        if app is not None:
            try:
                for idx in range(app.Documents.Count, 0, -1):
                    try: app.Documents(idx).Close(False)
                    except Exception: pass
            except Exception:
                pass
            try: app.Quit(False)
            except Exception: pass
        app = None
        gc.collect()
        if meta.get("owned_pid"):
            meta.update(_wait_or_terminate(handle, int(meta["owned_pid"])))
        elif handle:
            ctypes.windll.kernel32.CloseHandle(handle)
        _clear_ownership()
        pythoncom.CoUninitialize()


def _wait_for_calculation(excel, timeout: float = 120.0) -> None:
    deadline = time.monotonic() + timeout
    while int(excel.CalculationState) != XL_DONE:
        if time.monotonic() >= deadline:
            raise TimeoutError("Excel calculation did not reach xlDone before the timeout.")
        time.sleep(0.1)


def office_health() -> dict:
    result: dict = {"status": "PASS", "word": {}, "excel": {}}
    try:
        with word_session() as (word, meta):
            info = {"available": True, "version": str(word.Version), "owned_pid": meta.get("owned_pid")}
            legacy._probe_word_file_ops(word)
            info["file_operations"] = True
            result["word"] = info
    except Exception as exc:
        result["word"] = {"available": False, "error": str(exc)}
        result["status"] = "FAIL"
    try:
        with excel_session() as (excel, meta):
            info = {"available": True, "version": str(excel.Version), "owned_pid": meta.get("owned_pid")}
            legacy._probe_excel_file_ops(excel)
            info["file_operations"] = True
            result["excel"] = info
    except Exception as exc:
        result["excel"] = {"available": False, "error": str(exc)}
        result["status"] = "FAIL"
    return result


def _word_pdf(src: Path, pdf: Path) -> dict:
    with word_session() as (word, meta):
        doc = None
        try:
            doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
            legacy._require_opened(doc, src, "Word")
            doc.Repaginate()
            doc.ExportAsFixedFormat(str(pdf), 17)
            return {"status": "PASS", "engine": "word-com", "pdf": str(pdf), "session": meta}
        finally:
            if doc is not None:
                try: doc.Close(False)
                except Exception: pass


def _excel_pdf(src: Path, pdf: Path) -> dict:
    with excel_session() as (excel, meta):
        wb = None
        try:
            wb = excel.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
            legacy._require_opened(wb, src, "Excel")
            wb.ExportAsFixedFormat(0, str(pdf))
            return {"status": "PASS", "engine": "excel-com", "pdf": str(pdf), "session": meta}
        finally:
            if wb is not None:
                try: wb.Close(False)
                except Exception: pass


def render(src: Path, out: Path) -> dict:
    src = src.resolve(); out = out.resolve(); out.mkdir(parents=True, exist_ok=True)
    pdf = out / f"{src.stem}.pdf"
    suffix = src.suffix.lower()
    if suffix in {".docx", ".docm", ".doc"}:
        return _word_pdf(src, pdf)
    if suffix in {".xlsx", ".xlsm", ".xlsb", ".xls"}:
        return _excel_pdf(src, pdf)
    return {"status": "FAIL", "message": f"Unsupported Office file: {src}"}


def recalculate_xlsx(src: Path) -> dict:
    src = src.resolve()
    with excel_session() as (excel, meta):
        wb = None
        try:
            wb = excel.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
            legacy._require_opened(wb, src, "Excel")
            calc = {
                "calculation_mode": int(excel.Calculation),
                "iteration": bool(excel.Iteration),
                "max_iterations": int(excel.MaxIterations),
                "max_change": float(excel.MaxChange),
            }
            excel.CalculateFullRebuild()
            _wait_for_calculation(excel)
            wb.Save()
            calc["calculation_state"] = int(excel.CalculationState)
            return {"status": "PASS", "engine": "excel-com", "file": str(src), "calculation": calc, "session": meta}
        finally:
            if wb is not None:
                try: wb.Close(False)
                except Exception: pass


def _counta(excel, rng) -> int:
    try:
        return int(excel.WorksheetFunction.CountA(rng))
    except Exception:
        return 1


def autofit_xlsx(src: Path) -> dict:
    """Native measurement with finance-model constraints: expand only, preserve hidden/spacer dimensions."""
    src = src.resolve()
    with excel_session() as (excel, meta):
        wb = None
        try:
            wb = excel.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
            legacy._require_opened(wb, src, "Excel")
            changes = []; skipped = []
            for ws in wb.Worksheets:
                used = ws.UsedRange
                r0 = int(used.Row); c0 = int(used.Column)
                r1 = r0 + int(used.Rows.Count) - 1; c1 = c0 + int(used.Columns.Count) - 1
                sheet = {"sheet": ws.Name, "columns_expanded": 0, "rows_expanded": 0}
                for col_idx in range(c0, c1 + 1):
                    col = ws.Columns(col_idx)
                    try:
                        if bool(col.Hidden):
                            continue
                        original = float(col.ColumnWidth or ws.StandardWidth)
                        if original <= 3.0:
                            continue
                        segment = ws.Range(ws.Cells(r0, col_idx), ws.Cells(r1, col_idx))
                        if _counta(excel, segment) == 0:
                            continue
                        col.AutoFit()
                        fitted = float(col.ColumnWidth)
                        target = max(original, fitted)
                        if original <= 45.0:
                            target = min(target, 45.0)
                        col.ColumnWidth = target
                        if target > original + 0.01:
                            sheet["columns_expanded"] += 1
                    except Exception as exc:
                        skipped.append(f"{ws.Name}!column {col_idx}: {exc}")
                for row_idx in range(r0, r1 + 1):
                    row = ws.Rows(row_idx)
                    try:
                        if bool(row.Hidden):
                            continue
                        original = float(row.RowHeight or ws.StandardHeight)
                        row.AutoFit()
                        fitted = float(row.RowHeight or original)
                        target = max(original, fitted)
                        if original <= 72.0:
                            target = min(target, 72.0)
                        row.RowHeight = target
                        if target > original + 0.01:
                            sheet["rows_expanded"] += 1
                    except Exception as exc:
                        skipped.append(f"{ws.Name}!row {row_idx}: {exc}")
                changes.append(sheet)
            wb.Save()
            return {
                "status": "PASS", "engine": "excel-com", "file": str(src), "sheets": changes,
                "skipped": skipped[:100],
                "policy": "expand-only; preserve hidden/spacer dimensions; cap only automatic expansions",
                "session": meta,
            }
        finally:
            if wb is not None:
                try: wb.Close(False)
                except Exception: pass


def repaginate_docx(src: Path) -> dict:
    src = src.resolve()
    with word_session() as (word, meta):
        doc = None
        try:
            doc = word.Documents.Open(str(src), ReadOnly=False, AddToRecentFiles=False, ConfirmConversions=False)
            legacy._require_opened(doc, src, "Word")
            try: doc.Fields.Update()
            except Exception: pass
            try:
                for toc in doc.TablesOfContents: toc.Update()
            except Exception: pass
            doc.Repaginate(); doc.Save()
            return {"status": "PASS", "engine": "word-com", "file": str(src), "session": meta}
        finally:
            if doc is not None:
                try: doc.Close(False)
                except Exception: pass


def snapshot_xlsx(src: Path) -> dict:
    src = src.resolve()
    with excel_session() as (excel, meta):
        wb = None
        try:
            wb = excel.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True)
            legacy._require_opened(wb, src, "Excel")
            sheets = []
            for ws in wb.Worksheets:
                used = ws.UsedRange
                sheets.append({
                    "name": ws.Name, "visible": int(ws.Visible), "used_range": used.Address,
                    "rows": int(used.Rows.Count), "columns": int(used.Columns.Count),
                    "page_orientation": int(ws.PageSetup.Orientation),
                    "zoom": ws.Application.ActiveWindow.Zoom if ws.Application.ActiveWindow else None,
                })
            return {"status": "PASS", "engine": "excel-com", "file": str(src), "sheets": sheets, "session": meta}
        finally:
            if wb is not None:
                try: wb.Close(False)
                except Exception: pass


def snapshot_docx(src: Path) -> dict:
    src = src.resolve()
    with word_session() as (word, meta):
        doc = None
        try:
            doc = word.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
            legacy._require_opened(doc, src, "Word")
            doc.Repaginate(); sections = []
            for sec in doc.Sections:
                p = sec.PageSetup
                sections.append({
                    "orientation": int(p.Orientation), "page_width": float(p.PageWidth), "page_height": float(p.PageHeight),
                    "top_margin": float(p.TopMargin), "bottom_margin": float(p.BottomMargin),
                    "left_margin": float(p.LeftMargin), "right_margin": float(p.RightMargin),
                })
            return {
                "status": "PASS", "engine": "word-com", "file": str(src),
                "pages": int(doc.ComputeStatistics(2)), "paragraphs": int(doc.Paragraphs.Count),
                "tables": int(doc.Tables.Count), "sections": sections, "session": meta,
            }
        finally:
            if doc is not None:
                try: doc.Close(False)
                except Exception: pass


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Hardened native Microsoft Office COM worker")
    sub = parser.add_subparsers(dest="cmd", required=True)
    sub.add_parser("health")
    p = sub.add_parser("render"); p.add_argument("--input", required=True); p.add_argument("--out", required=True)
    for name in ("recalculate-xlsx", "autofit-xlsx", "repaginate-docx", "snapshot"):
        p = sub.add_parser(name); p.add_argument("--input", required=True)
    args = parser.parse_args(argv)
    try:
        if args.cmd == "health": data = office_health()
        elif args.cmd == "render": data = render(Path(args.input), Path(args.out))
        elif args.cmd == "recalculate-xlsx": data = recalculate_xlsx(Path(args.input))
        elif args.cmd == "autofit-xlsx": data = autofit_xlsx(Path(args.input))
        elif args.cmd == "repaginate-docx": data = repaginate_docx(Path(args.input))
        elif args.cmd == "snapshot":
            src = Path(args.input)
            data = snapshot_xlsx(src) if src.suffix.lower() in {".xlsx", ".xlsm", ".xlsb", ".xls"} else snapshot_docx(src)
        else: raise AssertionError(args.cmd)
    except Exception as exc:
        data = {"status": "FAIL", "error": type(exc).__name__, "message": str(exc)}
    return legacy.emit(data, 0 if data.get("status") == "PASS" else 2)


if __name__ == "__main__":
    raise SystemExit(main())
