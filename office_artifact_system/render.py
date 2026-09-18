from __future__ import annotations

from pathlib import Path
import shutil
import subprocess

from .com_bridge import com_render


def _rasterize(pdf: Path, out: Path, outputs: list[str]) -> None:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm or not pdf.exists():
        return
    prefix = out / "page"
    proc = subprocess.run(
        [pdftoppm, "-png", "-r", "120", str(pdf), str(prefix)],
        text=True,
        capture_output=True,
        check=False,
    )
    if proc.returncode == 0:
        outputs.extend(str(x) for x in sorted(out.glob("page-*.png")))


def _render_libreoffice(path: Path, out: Path) -> dict:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return {
            "status": "UNAVAILABLE",
            "engine": "libreoffice",
            "message": "LibreOffice/soffice not found on PATH.",
            "outputs": [],
        }
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out), str(path)]
    proc = subprocess.run(cmd, text=True, capture_output=True, check=False)
    pdf = out / f"{path.stem}.pdf"
    outputs = [str(pdf)] if pdf.exists() else []
    if proc.returncode != 0 or not pdf.exists():
        return {
            "status": "FAIL",
            "engine": "libreoffice",
            "command": cmd,
            "stdout": proc.stdout,
            "stderr": proc.stderr,
            "outputs": outputs,
        }
    _rasterize(pdf, out, outputs)
    return {
        "status": "PASS",
        "engine": "libreoffice",
        "command": cmd,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
        "outputs": outputs,
    }


def _render_com(path: Path, out: Path) -> dict:
    data = com_render(path, out)
    pdf_value = data.get("pdf")
    outputs = []
    if pdf_value:
        pdf = Path(pdf_value)
        local_pdf = out / f"{path.stem}.pdf"
        pdf = local_pdf if local_pdf.exists() else pdf
        if pdf.exists():
            outputs.append(str(pdf))
            _rasterize(pdf, out, outputs)
    data["outputs"] = outputs
    return data


def render_office(path: str | Path, out_dir: str | Path, engine: str = "auto") -> dict:
    path = Path(path).resolve()
    out = Path(out_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    if engine not in {"auto", "com", "libreoffice"}:
        raise ValueError(f"Unknown render engine: {engine}")

    if engine in {"auto", "com"}:
        com_result = _render_com(path, out)
        if com_result.get("status") == "PASS":
            return com_result
        if engine == "com":
            return com_result
    else:
        com_result = None

    fallback = _render_libreoffice(path, out)
    if engine == "auto" and com_result is not None:
        fallback["preferred_engine_result"] = com_result
    return fallback
