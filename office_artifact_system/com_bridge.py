from __future__ import annotations

import json
import os
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


def _windows_path(path: str | Path) -> str:
    p = Path(path).resolve()
    if os.name == "nt":
        return str(p)
    if _is_wsl():
        proc = subprocess.run(
            ["wslpath", "-w", str(p)],
            text=True,
            capture_output=True,
            check=False,
        )
        if proc.returncode == 0 and proc.stdout.strip():
            return proc.stdout.strip()
    return str(p)


def _candidate_works(cmd: list[str]) -> bool:
    try:
        proc = subprocess.run(
            [*cmd, "-c", "import pythoncom; import win32com.client"],
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
    # Windows launcher at its fixed absolute WSL path (py.exe is always
    # C:\Windows\py.exe when installed). Try default tag first, then pinned
    # versions because a stale default tag (e.g. missing C:\Python313) is
    # common after a Store/manual Python upgrade.
    launcher = Path("/mnt/c/Windows/py.exe")
    if launcher.exists():
        for args in (["-3"], ["-V:3.12"], ["-V:3.11"], []):
            candidates.append([str(launcher), *args])
    # Common per-user and machine-wide CPython locations. Glob instead of
    # hardcoding a username so no personal path ships in source.
    for pattern in (
        "/mnt/c/Users/*/AppData/Local/Programs/Python/Python*/python.exe",
        "/mnt/c/Program Files/Python*/python.exe",
        "/mnt/c/Python*/python.exe",
    ):
        try:
            import glob as _glob

            for match in sorted(_glob.glob(pattern)):
                if Path(match).exists():
                    candidates.append([match])
        except Exception:
            continue
    # Deduplicate while preserving order.
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
        exe = Path(configured)
        if exe.exists() or shutil.which(configured):
            return [configured], None
        return None, f"OFFICE_COM_PYTHON does not exist: {configured}"

    py = shutil.which("py.exe")
    if py:
        # Prefer an explicitly working launcher tag; the default "-3" tag
        # can point at a removed interpreter after upgrades.
        for args in (["-3"], ["-V:3.12"], ["-V:3.11"], []):
            cmd = [py, *args]
            if _candidate_works(cmd):
                return cmd, None
        return [py, "-3"], None

    python_exe = shutil.which("python.exe")
    if python_exe:
        return [python_exe], None

    # WSL: Windows directories are usually not on PATH, so probe the fixed
    # absolute locations before giving up. Each candidate is verified with
    # `--version` so a stale launcher tag never causes a cryptic failure.
    if _is_wsl() or Path("/mnt/c/Windows").exists():
        for cmd in _wsl_windows_python_candidates():
            if _candidate_works(cmd):
                return cmd, None

    if os.name == "nt":
        py = shutil.which("py")
        if py:
            return [py, "-3"], None
        python_exe = shutil.which("python")
        if python_exe:
            return [python_exe], None

    return None, "Windows Python was not found. Set OFFICE_COM_PYTHON or run windows/bootstrap.ps1."


def _invoke(action: str, *, input_path=None, out_dir=None) -> dict:
    prefix, error = _windows_python()
    if not prefix:
        return {"status": "UNAVAILABLE", "engine": "com", "message": error}

    # Route through the desktop transport: in a non-interactive session
    # (WSL => Session 0) Office COM blocks file operations, so the worker is
    # re-entered in the signed-in user's interactive session. On a native
    # interactive desktop the transport executes the worker directly.
    transport = _windows_path(PROJECT_ROOT / "windows" / "desktop_transport.py")
    cmd = [*prefix, transport, action]
    if input_path is not None:
        cmd += ["--input", _windows_path(input_path)]
    if out_dir is not None:
        cmd += ["--out", _windows_path(out_dir)]

    try:
        proc = subprocess.run(
            cmd, text=True, capture_output=True, check=False, timeout=360
        )
    except subprocess.TimeoutExpired:
        return {
            "status": "FAIL",
            "engine": "com",
            "message": "Windows COM worker timed out after 360s; Office may be showing a blocking dialog.",
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
