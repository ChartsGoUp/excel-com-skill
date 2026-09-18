# Architecture

## Objective

Turn Codex from a prompt-driven Office-file editor into a controlled artifact production system that uses native Microsoft Office whenever it materially improves fidelity.

## Core principle

Use each layer for what it does best:

- **Codex** for reasoning, routing, planning, and judgment.
- **Python / OOXML** for deterministic inspection, comparison, validation, and regression testing.
- **Microsoft Word / Excel COM** for preferred native editing, formatting, calculation, pagination, and PDF export when Office is available on Windows.
- **LibreOffice** as the cross-platform fallback renderer and compatibility check.
- **Visual QA** as the final layout gate.

## Layers

```text
User request
   ↓
AGENTS.md + constitution
   ↓
Office router skill
   ↓
Artifact skill + domain skill
   ↓
Rules + templates + golden examples
   ↓
Python / OOXML inspection
   ↓
Native Office mutation layer
   ├─ Word COM
   └─ Excel COM
   ↓
Candidate DOCX/XLSX
   ↓
Deterministic validators
   ↓
Source-fidelity comparison
   ↓
Native Office PDF export
   ↓
LibreOffice fallback / compatibility render
   ↓
Visual QA
   ↓
QA report
   ↓
Versioned deliverable
```

## 1. Instructions layer

`AGENTS.md` is the Codex entrypoint. `CLAUDE.md` is the project constitution. The constitution defines non-negotiable behavior and the mandatory production loop.

## 2. Skill layer

Repository skills are stored under `.agents/skills/<name>/SKILL.md`.

The router uses progressive disclosure. Codex sees concise skill descriptions first. It loads full skill instructions only when the task matches.

## 3. Rules layer

The `rules/` directory stores stable standards that should not be re-inferred from examples on every task.

Rule precedence inside the project:

1. current user instruction,
2. current-file-specific constraint,
3. custom rules,
4. artifact rules,
5. domain rules,
6. approved template,
7. golden example,
8. general professional judgment.

## 4. Reference layer

`source-data/golden-examples/` contains immutable examples supplied by the owner.

Golden examples answer: "What does excellent look like?"

Templates answer: "What exact reusable structure should this artifact start from?"

The two concepts are intentionally separate.

## 5. Inspection and validation layer

`python-docx`, `openpyxl`, `lxml`, ZIP/XML inspection, and the project validators are the preferred tools for:

- content extraction,
- formula and value comparison,
- style inventory,
- source-vs-candidate preservation checks,
- package validation,
- regression tests,
- workbook/document structure inspection.

Python/OOXML is the authoritative **verification** layer even when COM is used to edit the file.

## 6. Native Office mutation layer

When Microsoft Office is installed and COM automation is available, Word and Excel are the preferred engines for operations where native rendering matters.

### Word COM is preferred for

- final formatting and style application,
- table AutoFit and width behavior,
- list/numbering operations that depend on Word,
- section/page setup,
- page breaks and pagination,
- field updates,
- TOC updates,
- repagination,
- native PDF export.

### Excel COM is preferred for

- final formatting,
- AutoFit row/column sizing,
- workbook calculation and `CalculateFullRebuild`,
- charts and chart layout,
- conditional formatting and native features,
- print areas and page setup,
- freeze panes and window state,
- pivot/native workbook features,
- native PDF export.

COM is never allowed to overwrite immutable source files. Work on a candidate copy.

## 7. WSL-to-Windows bridge

Codex may run in WSL while Word and Excel run on the Windows host.

The bridge is:

```text
Codex / Linux Python in WSL
        ↓
office_artifact_system.com_bridge
        ↓
Windows Python launcher (`py.exe` / configured python.exe)
        ↓
windows/office_worker.py
        ↓
pywin32 COM
        ↓
WINWORD.EXE / EXCEL.EXE
```

The bridge converts WSL paths to Windows paths with `wslpath` when needed.

The Windows runtime is intentionally separate from the Linux virtual environment because `pywin32` is Windows-only.

## 8. Rendering layer

Rendering priority:

1. **Microsoft Office COM → PDF** when COM is available.
2. **LibreOffice → PDF** as fallback.
3. `pdftoppm` converts PDF pages to PNG when available.

The `render` command uses `--engine auto` by default. `auto` tries COM first and falls back to LibreOffice.

Rendered images are evidence for visual inspection. Automated structural checks do not replace visual QA.

## 9. Editing vs validation priority

### Editing / formatting fidelity

For Excel:
1. Excel COM
2. `openpyxl`
3. direct OOXML
4. LibreOffice fallback

For Word:
1. Word COM
2. `python-docx`
3. direct OOXML
4. LibreOffice fallback

### Validation

1. Python / OOXML deterministic checks
2. COM inspection where native state matters
3. visual inspection

## 10. Evidence layer

Each substantial task writes a QA report under `analysis/`. Issued files go to `deliverables/` only after the required checks pass.

## 11. Multi-agent layer

Optional reviewer profiles live under `.codex/agents/`. The primary workflow does not depend on custom-agent support because Codex versions can differ in how named agents are exposed. Skills and deterministic tools remain the authoritative workflow.

When custom reviewers are available:

- `docx-reviewer`: read-only Word layout/content review.
- `xlsx-reviewer`: read-only workbook/formula review.
- `qa-verifier`: read-only adversarial final verification.

## 12. Expansion path

Add new artifact families by creating:

1. a rule file,
2. a specialized skill,
3. an approved template or golden example,
4. deterministic validator checks where the rules are machine-testable,
5. native Office helpers where Office behavior matters,
6. regression fixtures.

Do not enlarge a general skill until it becomes ambiguous. Prefer small specialized skills with explicit trigger descriptions.
