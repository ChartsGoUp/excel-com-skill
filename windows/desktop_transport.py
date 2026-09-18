"""Run the Office worker in the signed-in user's interactive Windows session."""
from __future__ import annotations

import ctypes
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from contextlib import contextmanager


MUTATING_ACTIONS = {"recalculate-xlsx", "autofit-xlsx", "repaginate-docx"}


PROCESS_TERMINATE = 0x0001
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
SYNCHRONIZE = 0x00100000


class FILETIME(ctypes.Structure):
    _fields_ = [("dwLowDateTime", ctypes.c_uint32), ("dwHighDateTime", ctypes.c_uint32)]


def _creation_filetime(handle) -> int | None:
    kernel32 = ctypes.windll.kernel32
    creation, exit_time, kernel_time, user_time = FILETIME(), FILETIME(), FILETIME(), FILETIME()
    if not kernel32.GetProcessTimes(
        handle,
        ctypes.byref(creation),
        ctypes.byref(exit_time),
        ctypes.byref(kernel_time),
        ctypes.byref(user_time),
    ):
        return None
    return (int(creation.dwHighDateTime) << 32) | int(creation.dwLowDateTime)


def _terminate_verified_owned_process(record: dict) -> dict:
    """Terminate only the exact Office process instance published by the worker."""
    pid = int(record.get("pid") or 0)
    expected = record.get("creation_filetime")
    if not pid or expected is None:
        return {"watchdog_cleanup": "ownership record incomplete; no process terminated"}
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(
        SYNCHRONIZE | PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION,
        False,
        pid,
    )
    if not handle:
        return {"watchdog_cleanup": "owned Office process already exited"}
    try:
        actual = _creation_filetime(handle)
        if actual != int(expected):
            return {
                "watchdog_cleanup": (
                    "PID exists but process creation identity changed; refusing termination"
                )
            }
        if kernel32.TerminateProcess(handle, 1):
            kernel32.WaitForSingleObject(handle, 5000)
            return {
                "watchdog_cleanup": "terminated verified owned Office process",
                "owned_office_pid": pid,
            }
        return {
            "watchdog_cleanup": "verified owned Office process could not be terminated",
            "owned_office_pid": pid,
        }
    finally:
        kernel32.CloseHandle(handle)


def session_id() -> int:
    value = ctypes.c_ulong()
    if not ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(), ctypes.byref(value)):
        raise ctypes.WinError()
    return value.value


def _is_wsl_unc(value: str) -> bool:
    lowered = value.lower()
    return lowered.startswith("\\\\wsl.localhost\\") or lowered.startswith("\\\\wsl$\\")


def _arg_value(args: list[str], flag: str) -> tuple[int, str] | None:
    try:
        idx = args.index(flag)
    except ValueError:
        return None
    if idx + 1 >= len(args):
        return None
    return idx + 1, args[idx + 1]


def _atomic_copy_back(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_name(f".{dst.name}.office-artifact-{uuid.uuid4().hex}.tmp")
    try:
        shutil.copy2(src, tmp)
        os.replace(tmp, dst)
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


@contextmanager
def _mutation_transport_lock(args: list[str], timeout: float = 90.0):
    if not args or args[0] not in MUTATING_ACTIONS:
        yield
        return
    input_arg = _arg_value(args, "--input")
    if not input_arg:
        yield
        return
    import msvcrt
    canonical = os.path.normcase(os.path.abspath(input_arg[1]))
    key = hashlib.sha256(canonical.encode("utf-8", errors="surrogatepass")).hexdigest()
    root = Path(tempfile.gettempdir()) / "office-artifact-transport-locks"
    root.mkdir(parents=True, exist_ok=True)
    fh = open(root / f"{key}.lock", "a+b")
    acquired = False
    try:
        fh.seek(0, os.SEEK_END)
        if fh.tell() == 0:
            fh.write(b"0"); fh.flush()
        deadline = time.monotonic() + timeout
        while True:
            try:
                fh.seek(0)
                msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                acquired = True
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise TimeoutError(f"Timed out waiting for workbook mutation lock: {input_arg[1]}")
                time.sleep(0.1)
        yield
    finally:
        if acquired:
            try:
                fh.seek(0); msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
            except OSError:
                pass
        fh.close()


def _prepare_staging(args: list[str], root: Path) -> tuple[list[str], dict]:
    """Stage WSL UNC files onto local NTFS before Office opens them."""
    staged = list(args)
    state: dict = {
        "input_original": None,
        "input_staged": None,
        "output_original": None,
        "output_staged": None,
        "mutating": bool(args and args[0] in MUTATING_ACTIONS),
    }

    input_arg = _arg_value(staged, "--input")
    if input_arg and _is_wsl_unc(input_arg[1]):
        idx, original = input_arg
        source = Path(original)
        local_input = root / "input" / source.name
        local_input.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, local_input)
        staged[idx] = str(local_input)
        state["input_original"] = source
        state["input_staged"] = local_input

    out_arg = _arg_value(staged, "--out")
    if out_arg and _is_wsl_unc(out_arg[1]):
        idx, original = out_arg
        local_out = root / "output"
        local_out.mkdir(parents=True, exist_ok=True)
        staged[idx] = str(local_out)
        state["output_original"] = Path(original)
        state["output_staged"] = local_out

    return staged, state


def _finalize_staging(result: dict, state: dict) -> dict:
    if result.get("status") == "PASS":
        if state.get("mutating") and state.get("input_original") and state.get("input_staged"):
            _atomic_copy_back(Path(state["input_staged"]), Path(state["input_original"]))
            if result.get("file"):
                result["file"] = str(state["input_original"])

        if state.get("output_original") and state.get("output_staged"):
            original = Path(state["output_original"])
            original.mkdir(parents=True, exist_ok=True)
            for child in Path(state["output_staged"]).iterdir():
                if child.is_file():
                    _atomic_copy_back(child, original / child.name)
            if result.get("pdf"):
                result["pdf"] = str(original / Path(result["pdf"]).name)

    if state.get("input_staged") or state.get("output_staged"):
        result["staging"] = {
            "used_local_ntfs": True,
            "input_staged": bool(state.get("input_staged")),
            "output_staged": bool(state.get("output_staged")),
        }
    return result


def run_worker(args: list[str]) -> dict:
    with tempfile.TemporaryDirectory(prefix="office-com-watchdog-") as directory:
        ownership = Path(directory) / "ownership.json"
        env = os.environ.copy()
        env["OFFICE_COM_OWNERSHIP_FILE"] = str(ownership)
        try:
            proc = subprocess.run(
                [sys.executable, str(Path(__file__).with_name("office_worker.py")), *args],
                capture_output=True,
                text=True,
                timeout=270,
                env=env,
            )
            try:
                data = json.loads(proc.stdout)
            except (ValueError, TypeError):
                data = {"status": "FAIL", "message": "Office worker did not return JSON."}
            data["returncode"] = proc.returncode
            if proc.stderr.strip():
                data["stderr"] = proc.stderr.strip()
            data["session_id"] = session_id()
            return data
        except subprocess.TimeoutExpired:
            cleanup = {}
            if ownership.exists():
                try:
                    cleanup = _terminate_verified_owned_process(
                        json.loads(ownership.read_text(encoding="utf-8"))
                    )
                except Exception as exc:
                    cleanup = {"watchdog_cleanup": f"cleanup failed: {exc}"}
            return {
                "status": "FAIL",
                "message": (
                    "Office worker timed out. The watchdog attempted cleanup only "
                    "for the exact Office process identity published by this worker."
                ),
                **cleanup,
            }


def desktop_job(request: Path) -> int:
    boot = {"python": sys.executable, "argv": sys.argv[1:]}
    try:
        boot["session"] = session_id()
        payload = json.loads(request.read_text(encoding="utf-8"))
        result = run_worker(payload["args"])
    except Exception as exc:
        result = {"status": "FAIL", "message": repr(exc)}
    result["transport"] = "interactive-scheduled-task"
    result["boot"] = boot
    target = request.with_name("result.json")
    temporary = target.with_suffix(".tmp")
    temporary.write_text(json.dumps(result), encoding="utf-8")
    temporary.replace(target)
    return 0 if result.get("status") == "PASS" else 2


def _run_interactive(args: list[str], request_root: Path) -> dict:
    request = request_root / "request.json"
    request.write_text(json.dumps({"args": args}), encoding="utf-8")
    powershell = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    proc = subprocess.run(
        [
            str(powershell),
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(Path(__file__).with_name("run_desktop.ps1")),
            "-Python",
            sys.executable,
            "-Runner",
            str(Path(__file__).resolve()),
            "-Request",
            str(request),
        ],
        capture_output=True,
        text=True,
        timeout=330,
    )
    result_file = request.with_name("result.json")
    if not result_file.exists():
        return {
            "status": "FAIL",
            "transport": "interactive-scheduled-task",
            "message": (
                "Desktop transport failed. Keep the same Windows user signed in "
                "and check Task Scheduler permissions."
            ),
            "stderr": proc.stderr.strip(),
            "returncode": proc.returncode,
        }
    return json.loads(result_file.read_text(encoding="utf-8"))


def invoke(args: list[str]) -> dict:
    # Serialize mutations against the ORIGINAL input path so two WSL jobs cannot
    # independently stage and then race during copy-back to the same workbook.
    with _mutation_transport_lock(args):
        # Always use a local NTFS staging directory. WSL UNC paths are copied into it
        # before Office sees them, then copied back atomically after successful writes.
        with tempfile.TemporaryDirectory(prefix="office-com-stage-") as stage_dir:
            stage_root = Path(stage_dir)
            staged_args, state = _prepare_staging(args, stage_root)
            if session_id() != 0:
                result = run_worker(staged_args)
                result["transport"] = "direct-interactive"
            else:
                # run_desktop.ps1 grants the signed-in user access to this directory.
                # Keep staged input/output and request/result in the same ACL scope.
                result = _run_interactive(staged_args, stage_root)
            return _finalize_staging(result, state)


def main() -> int:
    try:
        if len(sys.argv) == 3 and sys.argv[1] == "--job":
            return desktop_job(Path(sys.argv[2]))
        result = invoke(sys.argv[1:])
    except Exception as exc:
        result = {"status": "FAIL", "message": str(exc)}
    print(json.dumps(result, indent=2))
    return 0 if result.get("status") == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
