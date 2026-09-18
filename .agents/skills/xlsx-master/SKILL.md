---
name: xlsx-master
description: Create, edit, reformat, or repair professional Microsoft Excel XLSX workbooks. Use for models, schedules, trackers, reconciliations, statements, registers, and accounting workpapers. Do not use for Word documents.
---

# XLSX Master

Read:
- `../../../rules/02-xlsx-rules.md`
- `../../../rules/03-financial-modeling-rules.md` when the workbook is a financial model
- `../../../rules/05-qa-checklist.md`
- `../../../rules/06-reference-guide.md`
- `../../../rules/08-custom-rules.md`
- `../../../rules/09-golden-example-analysis.md`
- `../../../rules/10-native-office-com.md`
- `../../../config/financial-model-style.json` for finance-model semantic styles

Workflow:
1. Inspect the source with `python ../../../tools/office_artifacts.py inspect-xlsx <source>`.
2. Read `edit_preflight`. If advanced/native features are present, prefer Excel COM or surgical OOXML rather than a routine `openpyxl` save cycle.
3. Copy the source into a task working directory. Never edit `source-data/`.
4. Define the intended workbook/sheet/period structure before writing formulas.
5. Prefer Excel COM for native workbook objects, preservation-sensitive edits, final formatting, constrained AutoFit, chart/native feature handling, print setup, calculation, and native PDF export.
6. Use `openpyxl`, direct OOXML, and `office_artifact_system.xlsx_style` for deterministic low-risk structure edits, formula/value inspection, bulk changes, and semantic styling.
7. Preserve formulas, values, names, tables, validations, merged ranges, workbook metadata, and advanced OOXML features unless the task requires a change.
8. Use semantic styles and restrained professional formatting. Avoid full-grid borders and decorative dashboards.
9. When formula logic or cached results change, run `com-recalculate-xlsx` on the candidate when COM is available. Confirm calculation reaches `xlDone`.
10. Use `com-autofit-xlsx` only as constrained native measurement. It must not shrink deliberate model dimensions or destroy spacer columns. Review the result visually.
11. Run `validate-xlsx` on the recalculated candidate. Treat calculated-value errors as failures.
12. Run `compare-xlsx` against the source when preservation matters. Review value/formula, style, dimension, structure, and OOXML-feature differences.
13. Render important sheets with `render --engine auto`; this prefers Excel COM and falls back to LibreOffice.
14. Inspect hierarchy, widths, wrapping, assumptions, number formats, totals, sheet order, tab colors, freeze panes, gridlines, print layout, chart placement, and output readability.
15. For financial models, run statement/roll-forward/check logic from `03-financial-modeling-rules.md`.
16. Re-run failed checks and issue only the validated file.

Use the supplied Three Statement Financial Model as the primary golden example only for finance-model-style workbooks.

Never treat COM calculation as a substitute for source/formula/structure comparison.
