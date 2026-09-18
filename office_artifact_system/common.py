from __future__ import annotations

from dataclasses import dataclass, asdict
from hashlib import sha256
from pathlib import Path
from typing import Any
import json


@dataclass
class Finding:
    severity: str
    code: str
    message: str
    location: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def hash_file(path: str | Path) -> str:
    h = sha256()
    with Path(path).open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def write_json(data: Any, output: str | Path | None = None) -> None:
    text = json.dumps(data, indent=2, ensure_ascii=False, sort_keys=True)
    if output:
        Path(output).write_text(text + "\n", encoding="utf-8")
    else:
        print(text)
