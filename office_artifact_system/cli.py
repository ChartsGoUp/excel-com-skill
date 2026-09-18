from __future__ import annotations

import argparse

from .common import write_json
from .docx_ops import inspect_docx, validate_docx, compare_docx
from .xlsx_ops import inspect_xlsx, validate_xlsx, compare_xlsx
from .render import render_office
from .qa import qa_docx, qa_xlsx
from .com_bridge import (
    com_health,
    com_recalculate_xlsx,
    com_autofit_xlsx,
    com_repaginate_docx,
    com_snapshot,
)


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Office Artifact System CLI")
    sub = p.add_subparsers(dest="cmd", required=True)

    for name in ("inspect-docx", "validate-docx", "inspect-xlsx", "validate-xlsx"):
        s = sub.add_parser(name)
        s.add_argument("file")
        s.add_argument("--json", dest="json_out", default=None)

    for name in ("compare-docx", "compare-xlsx"):
        s = sub.add_parser(name)
        s.add_argument("source")
        s.add_argument("candidate")
        s.add_argument("--json", dest="json_out", default=None)

    for name in ("qa-docx", "qa-xlsx"):
        s = sub.add_parser(name)
        s.add_argument("source")
        s.add_argument("candidate")
        s.add_argument("--render-out", default=None)
        s.add_argument("--json", dest="json_out", default=None)

    s = sub.add_parser("render")
    s.add_argument("file")
    s.add_argument("--out", required=True)
    s.add_argument("--engine", choices=("auto", "com", "libreoffice"), default="auto")
    s.add_argument("--json", dest="json_out", default=None)

    s = sub.add_parser("engine-status")
    s.add_argument("--json", dest="json_out", default=None)

    for name in ("com-recalculate-xlsx", "com-autofit-xlsx", "com-repaginate-docx", "com-snapshot"):
        s = sub.add_parser(name)
        s.add_argument("file")
        s.add_argument("--json", dest="json_out", default=None)

    args = p.parse_args(argv)

    if args.cmd == "inspect-docx":
        data = inspect_docx(args.file)
    elif args.cmd == "validate-docx":
        data = validate_docx(args.file)
    elif args.cmd == "compare-docx":
        data = compare_docx(args.source, args.candidate)
    elif args.cmd == "inspect-xlsx":
        data = inspect_xlsx(args.file)
    elif args.cmd == "validate-xlsx":
        data = validate_xlsx(args.file)
    elif args.cmd == "compare-xlsx":
        data = compare_xlsx(args.source, args.candidate)
    elif args.cmd == "qa-docx":
        data = qa_docx(args.source, args.candidate, args.render_out)
    elif args.cmd == "qa-xlsx":
        data = qa_xlsx(args.source, args.candidate, args.render_out)
    elif args.cmd == "render":
        data = render_office(args.file, args.out, args.engine)
    elif args.cmd == "engine-status":
        data = com_health()
    elif args.cmd == "com-recalculate-xlsx":
        data = com_recalculate_xlsx(args.file)
    elif args.cmd == "com-autofit-xlsx":
        data = com_autofit_xlsx(args.file)
    elif args.cmd == "com-repaginate-docx":
        data = com_repaginate_docx(args.file)
    elif args.cmd == "com-snapshot":
        data = com_snapshot(args.file)
    else:
        raise AssertionError(args.cmd)

    write_json(data, getattr(args, "json_out", None))

    if data.get("status") == "FAIL":
        return 2
    if args.cmd.startswith("compare-") and not data.get("content_equal", False):
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
