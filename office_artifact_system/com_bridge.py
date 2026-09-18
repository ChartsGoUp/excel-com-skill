from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
WORKER = PROJECT_ROOT / "windows" / "office_worker.py"


def _is_wsl() -> bool:
    try:
        return "microsoft" in Path("/proc/version").read_text(errors="ignore").lower()
    except Exception:
        return False


def _run_wslpath(flag: str, value: str) -> str:
    proc = subprocess.run(
        ["wslpath", flag, value],
        text=True,
        capture_output=True,
        check=False,
        timeout=20,
    )
    if proc.returncode != 0 or not proc.stdout.strip():
        raise RuntimeError(
            f"wslpath {flag} failed for {value!r}: "
            f"{proc.stderr.strip() or 'no diagnostic returned'}"
        )
    return proc.stdout.strip()


def _windows_path(path: str | Path) -> str:
    p = Path(path).expanduser().resolve()
    if os.name == "nt":
        return str(p)
    if _is_wsl():
        # Fail closed. Passing /home/... through to Windows Office produces
        # misleading file-not-found/Trust Center failures.
        return _run_wslpath("-w", str(p))
    return str(p)


def _normalize_configured_windows_python(configured: str) -> str:
    configured = configured.strip().strip('"')
    if not _is_wsl():
        return configured
    # Accept either /mnt/c/... or a normal Windows C:\... value in WSL.
    if re.match(r"^[A-Za-z]:[\\/]", configured) or configured.startswith("\\\\"):
        return _run_wslpath("-u", configured)
    return configured


def _candidate_works(cmd: list[str]) -> bool:
    try:
        proc = subprocess.run(
            [*cmd, "-c", "import pythoncom; import win32com.client; import win32process"],
            text=True,
            capture_output=True,
            check=False,
            timeout=20,
        )
        return proc.returncode == 0
    except Exception:
        return False


def _wsl_windows_python_candidates() -> list[list[str]]:
    candidates: list[list[str]] = []
    launcher = Path("/mnt/c/Windows/py.exe")
    if launcher.exists():
        for args in (["-3"], ["-V:3.13"], ["-V:3.12"], ["-V:3.11"], []):
            candidates.append([str(launcher), *args])
    for pattern in (
        "/mnt/c/Users/*/AppData/Local/Programs/Python/Python*/python.exe",
        "/mnt/c/Program Files/Python*/python.exe",
        "/mnt/c/Python*/python.exe",
    ):
        try:
            import glob as _glob
            for match in sorted(_glob.glob(pattern), reverse=True):
                if Path(match).exists():
                    candidates.append([match])
        except Exception:
            continue
    seen: set[tuple[str, ...]] = set()
    unique: list[list[str]] = []
    for cmd in candidates:
        key = tuple(cmd)
        if key not in seen:
            seen.add(key)
            unique.append(cmd)
    return unique


def _windows_python() -> tuple[list[str] | None, str | None]:
    configured = os.environ.get("OFFICE_COM_PYTHON")
    if configured:
        try:
            normalized = _normalize_configured_windows_python(configured)
        except Exception as exc:
            return None, f"Invalid OFFICE_COM_PYTHON: {exc}"
        exe = Path(normalized)
        if exe.exists() or shutil.which(normalized):
            cmd = [normalized]
            if _candidate_works(cmd):
                return cmd, None
            return None, f"OFFICE_COM_PYTHON exists but pywin32 imports failed: {configured}"
        return None, f"OFFICE_COM_PYTHON does not exist: {configured}"

    py = shutil.which("py.exe")
    if py:
        for args in (["-3"], ["-V:3.13"], ["-V:3.12"], ["-V:3.11"], []):
            cmd = [py, *args]
            if _candidate_works(cmd):
                return cmd, None

    python_exe = shutil.which("python.exe")
    if python_exe and _candidate_works([python_exe]):
        return [python_exe], None

    if _is_wsl() or Path("/mnt/c/Windows").exists():
        for cmd in _wsl_windows_python_candidates():
            if _candidate_works(cmd):
                return cmd, None

    if os.name == "nt":
        for name in ("py", "python"):
            exe = shutil.which(name)
            if exe:
                cmd = [exe, "-3"] if name == "py" else [exe]
                if _candidate_works(cmd):
                    return cmd, None

    return None, "Windows Python with pywin32 was not found. Set OFFICE_COM_PYTHON or run windows/bootstrap.ps1."


def _invoke(action: str, *, input_path=None, out_dir=None) -> dict:
    prefix, error = _windows_python()
    if not prefix:
        return {"status": "UNAVAILABLE", "engine": "com", "message": error}

    try:
        transport = _windows_path(PROJECT_ROOT / "windows" / "desktop_transport.py")
        cmd = [*prefix, transport, action]
        if input_path is not None:
            cmd += ["--input", _windows_path(input_path)]
        if out_dir is not None:
            # The directory may not exist yet. Convert its resolved parent and
            # leaf through wslpath by creating it first on the Linux side.
            Path(out_dir).expanduser().mkdir(parents=True, exist_ok=True)
            cmd += ["--out", _windows_path(out_dir)]
    except Exception as exc:
        return {
            "status": "FAIL",
            "engine": "com",
            "message": f"WSL/Windows path preparation failed: {exc}",
        }

    try:
        proc = subprocess.run(
            cmd, text=True, capture_output=True, check=False, timeout=360
        )
    except subprocess.TimeoutExpired:
        return {
            "status": "FAIL",
            "engine": "com",
            "message": (
                "Windows COM transport timed out after 360s. "
                "The transport/worker watchdog owns cleanup of its Office process."
            ),
        }

    stdout = proc.stdout.strip()
    try:
        data = json.loads(stdout) if stdout else {}
    except json.JSONDecodeError:
        data = {}

    if not data:
        data = {
            "status": "FAIL",
            "engine": "com",
            "message": "Windows COM worker did not return JSON.",
        }

    data.setdefault("returncode", proc.returncode)
    if proc.stderr.strip():
        data.setdefault("stderr", proc.stderr.strip())
    return data


def com_health() -> dict:
    return _invoke("health")


def com_render(path: str | Path, out_dir: str | Path) -> dict:
    return _invoke("render", input_path=path, out_dir=out_dir)


def com_recalculate_xlsx(path: str | Path) -> dict:
    return _invoke("recalculate-xlsx", input_path=path)


def com_autofit_xlsx(path: str | Path) -> dict:
    return _invoke("autofit-xlsx", input_path=path)


def com_repaginate_docx(path: str | Path) -> dict:
    return _invoke("repaginate-docx", input_path=path)


def com_snapshot(path: str | Path) -> dict:
    return _invoke("snapshot", input_path=path)
