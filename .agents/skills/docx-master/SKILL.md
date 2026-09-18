---
name: docx-master
description: Create, edit, reformat, or repair professional Microsoft Word DOCX files. Use for manuals, reports, memoranda, policies, agreements, registers, and other Word artifacts. Do not use for spreadsheets.
---

# DOCX Master

Read:
- `../../../rules/01-docx-rules.md`
- `../../../rules/05-qa-checklist.md`
- `../../../rules/06-reference-guide.md`
- `../../../rules/08-custom-rules.md`
- `../../../rules/09-golden-example-analysis.md`
- `../../../rules/10-native-office-com.md`

Workflow:
1. Inspect the source with `python ../../../tools/office_artifacts.py inspect-docx <source>`.
2. Copy the source into a task working directory. Never edit `source-data/`.
3. Prefer Word COM for final native formatting, table AutoFit, list behavior, pagination, field/TOC updates, and native PDF export when COM is available.
4. Use `python-docx`, direct OOXML, and `office_artifact_system.docx_style` for deterministic structure changes and operations that do not require Word's layout engine.
5. Preserve substantive text on formatting-only tasks.
6. After layout edits, use `com-repaginate-docx` when COM is available.
7. Run `validate-docx` on the candidate.
8. Run `compare-docx` against the source when preservation is required.
9. Render with `render --engine auto`; this prefers Word COM and falls back to LibreOffice.
10. Inspect every rendered page. Fix clipped tables, bad page breaks, heading-only pages, fake bullets, excessive whitespace, inconsistent fonts, and unreadable content.
11. Re-run failed checks.
12. Save QA evidence under `analysis/` and issue the final file under `deliverables/`.

For manuals and long-form accounting documents, use the supplied Accounting Manual as the primary golden example unless a more specific approved template exists.

Never claim that Word-native layout was verified if COM was unavailable.
