"""Run the existing worker in the same user's interactive Windows session."""
from __future__ import annotations

import ctypes
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile


def session_id() -> int:
    value = ctypes.c_ulong()
    if not ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(), ctypes.byref(value)):
        raise ctypes.WinError()
    return value.value


def run_worker(args: list[str]) -> dict:
    try:
        proc = subprocess.run(
            [sys.executable, str(Path(__file__).with_name('office_worker.py')), *args],
            capture_output=True, text=True, timeout=270,
        )
        try:
            data = json.loads(proc.stdout)
        except (ValueError, TypeError):
            data = {'status': 'FAIL', 'message': 'Office worker did not return JSON.'}
        data['returncode'] = proc.returncode
        if proc.stderr.strip():
            data['stderr'] = proc.stderr.strip()
        data['session_id'] = session_id()
        return data
    except subprocess.TimeoutExpired:
        return {'status': 'FAIL', 'message': 'Office worker timed out; inspect automation-owned Office instances. No unrelated process was terminated.'}


def desktop_job(request: Path) -> int:
    boot = {'python': sys.executable, 'argv': sys.argv[1:]}
    try:
        boot['session'] = session_id()
        payload = json.loads(request.read_text(encoding='utf-8'))
        result = run_worker(payload['args'])
    except Exception as exc:  # never leave the scheduled task without a result
        result = {'status': 'FAIL', 'message': repr(exc)}
    result['transport'] = 'interactive-scheduled-task'
    result['boot'] = boot
    target = request.with_name('result.json')
    temporary = target.with_suffix('.tmp')
    temporary.write_text(json.dumps(result), encoding='utf-8')
    temporary.replace(target)
    return 0 if result.get('status') == 'PASS' else 2


def invoke(args: list[str]) -> dict:
    if session_id() != 0:
        result = run_worker(args)
        result['transport'] = 'direct-interactive'
        return result
    # A per-user temporary directory, never a shared fixed runner filename.
    with tempfile.TemporaryDirectory(prefix='office-com-') as directory:
        request = Path(directory) / 'request.json'
        request.write_text(json.dumps({'args': args}), encoding='utf-8')
        powershell = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
        proc = subprocess.run([
            str(powershell), '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-File', str(Path(__file__).with_name('run_desktop.ps1')),
            '-Python', sys.executable, '-Runner', str(Path(__file__).resolve()),
            '-Request', str(request),
        ], capture_output=True, text=True, timeout=330)
        result_file = request.with_name('result.json')
        if not result_file.exists():
            return {'status': 'FAIL', 'transport': 'interactive-scheduled-task',
                    'message': 'Desktop transport failed. Keep the same Windows user signed in and check Task Scheduler permissions.',
                    'stderr': proc.stderr.strip(), 'returncode': proc.returncode}
        return json.loads(result_file.read_text(encoding='utf-8'))


def main() -> int:
    try:
        if len(sys.argv) == 3 and sys.argv[1] == '--job':
            return desktop_job(Path(sys.argv[2]))
        result = invoke(sys.argv[1:])
    except Exception as exc:
        result = {'status': 'FAIL', 'message': str(exc)}
    print(json.dumps(result, indent=2))
    return 0 if result.get('status') == 'PASS' else 2


if __name__ == '__main__':
    raise SystemExit(main())
