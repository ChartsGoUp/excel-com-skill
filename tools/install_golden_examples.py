#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from office_artifact_system.common import hash_file

EXPECTED = {
    "docx": (
        "fe8fe3182c92036b9d6695dc534d372ffa55ae6666e7dd1a948a489e777c198a",
        "Golden - Accounting Manual.docx",
    ),
    "xlsx": (
        "7c6a1a00ef989a2601753eeaefc78ce261ce54ea5335bb089055f739011e187c",
        "Golden - Three Statement Financial Model.xlsx",
    ),
}


def install(kind: str, source: Path, target_dir: Path) -> None:
    expected, filename = EXPECTED[kind]
    actual = hash_file(source)
    if actual != expected:
        raise SystemExit(
            f"{kind} hash mismatch: expected {expected}, got {actual}. "
            "Refusing to install an unapproved golden example."
        )
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / filename
    shutil.copy2(source, target)
    print(f"installed {kind}: {target}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Install the approved owner-supplied golden examples.")
    parser.add_argument("--docx", type=Path, required=True)
    parser.add_argument("--xlsx", type=Path, required=True)
    args = parser.parse_args()
    target = ROOT / "source-data" / "golden-examples"
    install("docx", args.docx.resolve(), target)
    install("xlsx", args.xlsx.resolve(), target)
    print("golden example installation: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
