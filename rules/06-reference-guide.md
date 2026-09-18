# 06 — Reference Guide

This file tells the assistant which Project reference files to use and why.

Written rules override reference examples. Current user instructions override all standing references.

---

# Primary Golden Examples

## DOCX

| File | Artifact type | Use for | Do not copy |
|---|---|---|---|
| `source-data/golden-examples/Golden - Accounting Manual.docx` | Accounting manual / long-form professional report | Cover-page hierarchy, TOC, restrained corporate header, Part/Section hierarchy, dense professional page layout, table styling, pagination, lists | Avondale logo/name, Henderson Partners, source wording, company data, accounting conclusions |
| `09_GOLDEN_EXAMPLE_ANALYSIS.md` | Interpretation guide | Explains which visual/structural features of the DOCX example should be imitated | N/A |

### Default DOCX reference rule

For a manual, accounting report, policy manual, or similar long-form document, use the Accounting Manual as the
primary golden example unless:
1. the user provides another reference,
2. an approved task-specific template is available, or
3. the existing source file has a deliberate style that must be preserved.

---

## XLSX

| File | Artifact type | Use for | Do not copy |
|---|---|---|---|
| `source-data/golden-examples/Golden - Three Statement Financial Model.xlsx` | Financial model / investment banking | Workbook flow, assumptions, financial statements, section bars, input highlighting, historical/forecast columns, subtotal rules, debt schedules, DCF/LBO presentation, checks | Juice Co. data, assumptions, dates, valuation results, exact structure for unrelated workbooks |
| `09_GOLDEN_EXAMPLE_ANALYSIS.md` | Interpretation guide | Explains which features of the model should be reused and where | N/A |

### Default XLSX reference rule

For three-statement, DCF, LBO, valuation, debt, or other investment-banking-style workbooks, use the Three Statement
Financial Model as the primary golden example.

For ordinary accounting schedules, trackers, registers, or operational workbooks, use the written XLSX rules unless
a more appropriate approved workbook reference has been added.

---

# Template Registry

Add approved blank/semi-blank templates here as they become available.

## DOCX Templates

| File | Artifact type | Use for | Do not copy |
|---|---|---|---|
| `ADD_TEMPLATE_NAME.docx` | Example: Memorandum | Exact layout, margins, styles, header/footer | Entity names, dates, source-specific wording |

## XLSX Templates

| File | Artifact type | Use for | Do not copy |
|---|---|---|---|
| `ADD_TEMPLATE_NAME.xlsx` | Example: Reconciliation | Exact sheet structure, calculation areas, formatting | Source balances, company names |

---

# Reference precedence

1. Current user request
2. Current source-file constraints
3. `CUSTOM_RULES.md`
4. Written artifact/domain rules
5. Approved task-specific template
6. Primary golden example
7. Other golden examples
8. General professional judgment

---

# General reference-use rules

1. Select the closest matching artifact type.
2. Use templates for exact reusable structure.
3. Use golden examples for quality, judgment, density, and visual language.
4. Never copy source-specific facts from a golden example.
5. Do not combine incompatible aesthetics from several examples without a reason.
6. When a golden example conflicts with a written rule, follow the written rule.
7. Explain any deliberate major deviation when it materially affects the result.
