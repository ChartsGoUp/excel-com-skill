# Windows Office COM runtime

This directory contains the Windows-native automation worker used by WSL/Codex.

## Requirements

- Windows desktop Microsoft Office
- Python 3 for Windows
- `pywin32`

Install:

```powershell
powershell -ExecutionPolicy Bypass -File .\bootstrap.ps1
```

The worker is normally invoked through `office_artifact_system.com_bridge`. Direct usage is useful for diagnostics:

```powershell
py -3 .\office_worker.py health
py -3 .\office_worker.py render --input C:\path\file.docx --out C:\path\render
py -3 .\office_worker.py recalculate-xlsx --input C:\path\candidate.xlsx
py -3 .\office_worker.py autofit-xlsx --input C:\path\candidate.xlsx
py -3 .\office_worker.py repaginate-docx --input C:\path\candidate.docx
py -3 .\office_worker.py snapshot --input C:\path\file.xlsx
```

All mutating actions must target a working/candidate file, never an immutable source file.
