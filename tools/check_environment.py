#!/usr/bin/env python3
from __future__ import annotations

import argparse
import importlib.util
import json
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from office_artifact_system.com_bridge import com_health


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Check Office Artifact System dependencies")
    parser.add_argument(
        "--require-com",
        action="store_true",
        help="Fail unless native Word and Excel COM are available.",
    )
    args = parser.parse_args(argv)

    mods = {
        name: bool(importlib.util.find_spec(name))
        for name in ("docx", "openpyxl", "lxml", "PIL")
    }
    executables = {
        name: shutil.which(name)
        for name in ("soffice", "libreoffice", "pdftoppm", "pdfinfo", "py.exe", "python.exe")
    }
    com = com_health()

    data = {
        "python_modules": mods,
        "executables": executables,
        "office_com": com,
        "preferred_render_engine": "com" if com.get("status") == "PASS" else "libreoffice",
    }
    print(json.dumps(data, indent=2, sort_keys=True))

    basic_ok = mods["docx"] and mods["openpyxl"] and mods["lxml"]
    if not basic_ok:
        return 2
    if args.require_com and com.get("status") != "PASS":
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
