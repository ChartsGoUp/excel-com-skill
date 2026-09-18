#!/usr/bin/env python3
from __future__ import annotations

import json
import py_compile
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from office_artifact_system.common import hash_file


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(2)


def main() -> int:
    cfg_path = ROOT / "config" / "artifact-system.json"
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))

    engine_policy = cfg.get("engine_policy", {})
    if engine_policy.get("docx_editing", [None])[0] != "word-com":
        fail("DOCX engine policy must prefer word-com")
    if engine_policy.get("xlsx_editing", [None])[0] != "excel-com":
        fail("XLSX engine policy must prefer excel-com")
    if engine_policy.get("rendering", [None])[0] != "office-com":
        fail("rendering engine policy must prefer office-com")

    missing_optional = []
    for name, meta in cfg["golden_examples"].items():
        path = ROOT / meta["path"]
        if not path.is_file():
            if meta.get("required_in_repo", True):
                fail(f"missing golden example {name}: {path}")
            missing_optional.append(str(path))
            continue
        actual = hash_file(path)
        if actual != meta["sha256"]:
            fail(f"golden example hash mismatch for {name}: {actual}")

    for skill in cfg["skills"]:
        skill_file = ROOT / ".agents" / "skills" / skill / "SKILL.md"
        if not skill_file.is_file():
            fail(f"missing skill: {skill_file}")
        text = skill_file.read_text(encoding="utf-8")
        if not text.startswith("---\n") or f"name: {skill}" not in text:
            fail(f"invalid skill frontmatter: {skill_file}")

    required = [
        "rules/00-global-rules.md",
        "rules/01-docx-rules.md",
        "rules/02-xlsx-rules.md",
        "rules/05-qa-checklist.md",
        "rules/10-native-office-com.md",
        "office_artifact_system/cli.py",
        "office_artifact_system/com_bridge.py",
        "office_artifact_system/render.py",
        "tools/office_artifacts.py",
        "tools/check_environment.py",
        "windows/office_worker.py",
        "windows/bootstrap.ps1",
        "requirements-windows.txt",
    ]
    for rel in required:
        if not (ROOT / rel).is_file():
            fail(f"missing required file: {rel}")

    python_files = [
        ROOT / "office_artifact_system" / "cli.py",
        ROOT / "office_artifact_system" / "com_bridge.py",
        ROOT / "office_artifact_system" / "render.py",
        ROOT / "tools" / "check_environment.py",
        ROOT / "windows" / "office_worker.py",
    ]
    for path in python_files:
        try:
            py_compile.compile(str(path), doraise=True)
        except py_compile.PyCompileError as exc:
            fail(f"python syntax validation failed for {path}: {exc}")

    proc = subprocess.run(
        [sys.executable, str(ROOT / "tools" / "selftest.py")],
        text=True,
        capture_output=True,
    )
    if proc.returncode:
        print(proc.stdout)
        print(proc.stderr, file=sys.stderr)
        fail("selftest failed")

    print("project validation: PASS")
    print(proc.stdout.strip())
    for path in missing_optional:
        print(f"WARN: approved golden example not installed locally: {path}")
    print("INFO: native Office COM runtime is validated separately with tools/check_environment.py --require-com")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
