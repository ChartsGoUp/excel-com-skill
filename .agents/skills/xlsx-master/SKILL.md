---
name: xlsx-master
description: Create, edit, reformat, or repair professional Microsoft Excel XLSX workbooks. Use for models, schedules, trackers, reconciliations, statements, registers, and accounting workpapers. Do not use for Word documents.
---

# XLSX Master

Read:
- `../../../rules/02-xlsx-rules.md`
- `../../../rules/05-qa-checklist.md`
- `../../../rules/06-reference-guide.md`
- `../../../rules/08-custom-rules.md`
- `../../../rules/09-golden-example-analysis.md`
- `../../../rules/10-native-office-com.md`

Workflow:
1. Inspect the source with `python ../../../tools/office_artifacts.py inspect-xlsx <source>`.
2. Copy the source into a task working directory. Never edit `source-data/`.
3. Prefer Excel COM for final native formatting, AutoFit, chart/native feature handling, print setup, calculation, and native PDF export when COM is available.
4. Use `openpyxl`, direct OOXML, and `office_artifact_system.xlsx_style` for deterministic structure edits, formula/value inspection, and bulk changes where Excel's rendering engine is not required.
5. Preserve formulas, values, names, tables, validations, merged ranges, and workbook metadata unless the task requires a change.
6. Use restrained professional formatting. Avoid full-grid borders and decorative dashboards.
7. When formula logic or cached results change, run `com-recalculate-xlsx` on the candidate when COM is available.
8. Use `com-autofit-xlsx` only when native AutoFit is appropriate; review the result rather than accepting it blindly.
9. Run `validate-xlsx` on the candidate.
10. Run `compare-xlsx` against the source when the task is formatting-only or requires calculation preservation.
11. Render important sheets with `render --engine auto`; this prefers Excel COM and falls back to LibreOffice.
12. Inspect widths, wrapping, assumptions, number formats, totals, sheet order, tab colors, freeze panes, gridlines, print layout, and output readability.
13. Re-run failed checks and issue only the validated file.

Use the supplied Three Statement Financial Model as the primary golden example only for finance-model-style workbooks.

Never treat COM calculation as a substitute for source/formula comparison.
