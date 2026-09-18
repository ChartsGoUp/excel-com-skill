---
name: office-artifact-router
description: Route professional DOCX/XLSX creation, editing, remediation, or review tasks to the correct Office artifact skills and mandatory QA workflow. Use for any substantial Word or Excel artifact task in this project.
---

# Office Artifact Router

1. Read `../../../CLAUDE.md` and the rule precedence in `../../../rules/00-global-rules.md`.
2. Classify the artifact as DOCX, XLSX, or mixed.
3. Classify the intent as formatting-only, structural improvement, substantive edit, creation, review-only, or QA-only.
4. Load `$docx-master` for Word work.
5. Load `$xlsx-master` for Excel work.
6. Also load `$financial-modeling` for financial models, debt schedules, DCF, LBO, valuation, or investment-banking-style workbooks.
7. Also load `$accounting-artifacts` for accounting manuals, journals, reconciliations, policies, intercompany schedules, shareholder-loan work, or financial statements.
8. Read `../../../rules/10-native-office-com.md` for any editing, formatting, calculation, pagination, or rendering task.
9. Prefer Word/Excel COM for native mutation and rendering when COM is available. Use Python/OOXML as the deterministic inspection and validation layer.
10. Always load `$artifact-qa` before issue.
11. Never edit files in `source-data/`.
12. Write final validated outputs only to `deliverables/`.

For formatting-only work, require a source-vs-candidate content/formula comparison before issue.

If COM is unavailable, fall back to Python/OOXML editing and LibreOffice rendering, and record the fallback in QA evidence.
