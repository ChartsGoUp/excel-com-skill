# Office Artifact System

A Codex production system for professional Microsoft Word (`.docx`) and Excel (`.xlsx`) artifacts.

The system uses **native Microsoft Office COM as the preferred editing/rendering engine when available**, while retaining Python/OOXML for deterministic inspection and validation and LibreOffice as a fallback.

## What this project does

It combines:

- repository-scoped Codex skills,
- durable DOCX/XLSX rules,
- immutable golden examples,
- reusable Python formatting helpers,
- native Word/Excel COM automation,
- deterministic validators,
- source-vs-candidate comparison tools,
- native Office PDF export,
- LibreOffice/PDF fallback rendering,
- visual QA procedures,
- project-local tasks and memory,
- versioned deliverables.

## Start Codex

Run Codex from this project directory so the skills are in scope:

```bash
cd "projects/5. Office Artifact System"
codex
```

Then use a normal request or explicitly invoke a skill:

```text
$office-artifact-router Improve the attached workbook without changing formulas or values.
$docx-master Reformat this manual for readability and preserve all substantive text.
$xlsx-master Improve this schedule to a professional finance/accounting standard.
$artifact-qa Validate the candidate against the source and render it for final review.
```

## Architecture summary

```text
Codex reasoning
   ↓
Python / OOXML inspection
   ↓
Word COM / Excel COM mutation (preferred)
   ↓
Python / OOXML validation
   ↓
Office COM PDF export (preferred)
   ↓
LibreOffice fallback
   ↓
Visual QA
```

See `ARCHITECTURE.md` for the complete design.

## Directory map

```text
.agents/skills/          Codex skills
.codex/agents/           optional specialist reviewer profiles
rules/                   durable artifact and domain standards
source-data/             immutable originals and golden examples
templates/               approved reusable templates
examples/                human notes about reference use
office_artifact_system/  reusable Python library and WSL COM bridge
windows/                 Windows-native pywin32 worker and bootstrap
tools/                   command-line entrypoints
analysis/                QA reports and workpapers
deliverables/            validated issued artifacts
agent-workspace/         task bus, memory, decisions, ADRs
```

## Linux / WSL Python environment

Install the cross-platform dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Windows COM environment

The preferred production path requires:

- Windows
- desktop Microsoft Word and Excel
- Windows Python 3
- `pywin32`

From PowerShell:

```powershell
cd "<repo>\projects\5. Office Artifact System"
powershell -ExecutionPolicy Bypass -File .\windows\bootstrap.ps1
```

Or from WSL:

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File \
  "$(wslpath -w "$PWD/windows/bootstrap.ps1")"
```

The bootstrap installs `requirements-windows.txt` into Windows Python and runs a Word/Excel COM health check.

If the Windows Python launcher is not discoverable from WSL, set:

```bash
export OFFICE_COM_PYTHON="/mnt/c/Path/To/python.exe"
```

## Engine status

```bash
python tools/office_artifacts.py engine-status
```

Use `--require-com` when setting up the preferred production environment:

```bash
python tools/check_environment.py --require-com
```

## Rendering

`auto` is the default and prefers COM:

```bash
python tools/office_artifacts.py render file.docx --out analysis/render --engine auto
```

Force an engine when debugging:

```bash
python tools/office_artifacts.py render file.docx --out analysis/render --engine com
python tools/office_artifacts.py render file.docx --out analysis/render --engine libreoffice
```

## Native Office maintenance operations

```bash
python tools/office_artifacts.py com-recalculate-xlsx candidate.xlsx
python tools/office_artifacts.py com-autofit-xlsx candidate.xlsx
python tools/office_artifacts.py com-repaginate-docx candidate.docx
python tools/office_artifacts.py com-snapshot candidate.xlsx --json analysis/style.json
```

These commands must operate on a working/candidate copy, never an immutable source.

## Install the approved golden examples

The architecture registers the exact owner-supplied examples by SHA-256. Install them once per clone:

```bash
python tools/install_golden_examples.py \
  --docx /path/to/2.\ Accounting\ Manual\(1\).docx \
  --xlsx /path/to/13_three_statement_model_0_0.xlsx
```

If they are not installed, the written reference analysis remains usable, but native comparison against the binary originals is unavailable.

## Fallback rendering

For non-Windows environments or when Office COM is unavailable, install LibreOffice and optionally Poppler:

```bash
sudo apt-get update
sudo apt-get install -y libreoffice poppler-utils
```

## Validation

Check the environment:

```bash
python tools/check_environment.py
```

Run project validation and self-test:

```bash
python tools/validate_project.py
python tools/selftest.py
```

## Primary quality principle

The system separates five questions:

1. Did the artifact preserve the source where required?
2. Is the OOXML package structurally valid?
3. Are calculations and formulas intact?
4. Did native Office produce the expected result?
5. Does the rendered result look correct?

The file is ready only after the required gates pass.
