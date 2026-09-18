# Tools

Primary entrypoint:

```bash
python tools/office_artifacts.py --help
```

Environment:

```bash
python tools/check_environment.py
python tools/check_environment.py --require-com
```

Native Office:

```bash
python tools/office_artifacts.py engine-status
python tools/office_artifacts.py render <file> --out <dir> --engine auto
python tools/office_artifacts.py com-recalculate-xlsx <candidate.xlsx>
python tools/office_artifacts.py com-autofit-xlsx <candidate.xlsx>
python tools/office_artifacts.py com-repaginate-docx <candidate.docx>
python tools/office_artifacts.py com-snapshot <file>
```

`auto` prefers Word/Excel COM and falls back to LibreOffice.

Use `windows/bootstrap.ps1` to configure the separate Windows Python/pywin32 runtime.
