---
name: artifact-qa
description: Validate final DOCX/XLSX artifacts before issue. Use for deterministic structural checks, source-fidelity comparisons, formula/value checks, native Office rendering, visual inspection, and final release evidence.
---

# Artifact QA

Read:
- `../../../rules/05-qa-checklist.md`
- `../../../rules/10-native-office-com.md`

Required gates are task-dependent.

## DOCX
1. `validate-docx`
2. `compare-docx` when source preservation is required
3. Word COM repagination/field update when layout was materially changed and COM is available
4. `render --engine auto`
5. visual page inspection
6. LibreOffice compatibility render when interoperability matters

## XLSX
1. `validate-xlsx`
2. `compare-xlsx` when formulas/values must be preserved
3. Excel COM `CalculateFullRebuild` and save/reopen when calculation changes matter and COM is available
4. `render --engine auto`
5. visual inspection of important sheets/pages
6. LibreOffice compatibility render when interoperability matters

`auto` rendering prefers native Office COM and falls back to LibreOffice.

Write a concise QA record to `analysis/` with:
- commands,
- engine used,
- deterministic results,
- source-comparison result,
- render result,
- unresolved blockers.

Never convert a failed check into PASS by suppressing the error. Never claim a visual issue is fixed without inspecting the final render. Never state that native Office was verified when the task used only the fallback engine.
