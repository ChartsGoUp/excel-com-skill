# 02 — XLSX Rules

## 1. Objective

Workbooks should look like professional finance/accounting files, not generic dashboards.

Primary goals:
- easy to audit,
- easy to navigate,
- easy to understand,
- visually restrained,
- structurally consistent.

---

## 2. Default typography

Unless an approved template requires otherwise:

- Font: **Arial**
- Standard cell font: **10 pt**
- Small labels / dense schedules: **9 pt** where necessary
- Important headings: 10–12 pt, bold

Avoid mixed fonts.

---

## 3. Workbook structure

Use logical sheet order.

Typical order:
1. Cover / Read Me only when needed
2. Summary / Output
3. Inputs / Assumptions
4. Core calculations
5. Supporting schedules
6. Checks / Reconciliations
7. Raw / source data
8. Historical / archived support

Do not add a Read Me or Executive Summary sheet automatically.

---

## 4. Sheet naming

Sheet names must be:
- concise,
- descriptive,
- stable,
- easy to reference in formulas.

Avoid unnecessarily long names and ambiguous abbreviations.

---

## 5. Tab colors

Use a consistent restrained scheme by sheet role.

Suggested default:
- Summary / outputs: dark blue
- Inputs / assumptions: light blue
- Core schedules: neutral/dark grey
- Checks / controls: green
- Source/raw data: light grey
- Historical/archive: muted grey

If an approved workbook already uses a good scheme, preserve it.

---

## 6. Gridlines

For professionally formatted sheets:
- hide gridlines where the formatting provides sufficient structure.

Do not hide gridlines on raw-data sheets if they materially improve usability.

---

## 7. Inputs and assumptions

Inputs should be clearly identifiable.

Requirements:
- descriptive line-item label,
- clear value,
- unit where necessary,
- source or note where useful,
- no clipped labels,
- no awkward wrapped labels caused by narrow columns.

Do not create an executive-summary-style input box when a clean assumptions table is more appropriate.

---

## 8. Formulas

Requirements:
- preserve formulas unless changes are requested,
- avoid unexplained hardcodes inside formulas,
- avoid broken links,
- avoid formula inconsistencies within equivalent rows/columns,
- avoid hardcoding output values where a formula should exist,
- do not hide errors using broad `IFERROR(...,0)` unless zero is economically correct.

Check for:
- `#REF!`
- `#VALUE!`
- `#DIV/0!`
- `#NAME?`
- `#N/A`
- missing formulas
- external links
- inconsistent formulas

---

## 9. Investment-banking-style formula colors

Where appropriate and not overridden by a template:

- Hardcoded inputs: blue font
- Formulas: black font
- Links to other sheets in the same workbook: black or template standard
- Links to external workbooks: green font if intentionally retained
- Warnings / unresolved exceptions: red only when necessary

Do not use color as the only control.

---

## 10. Number formats

Use professional financial formats.

Default principles:
- negatives in parentheses,
- zeros shown as `-` where appropriate,
- consistent decimal places,
- percentages clearly formatted,
- dates consistently formatted,
- currencies and units clearly identified,
- large values use appropriate `$`, `000s`, `mm`, or other disclosed unit conventions.

Do not mix inconsistent formats within the same schedule.

---

## 11. Borders

Use borders sparingly.

Prefer:
- section spacing,
- bold,
- fill,
- alignment,
- subtotal rules,

before adding boxes around every cell.

Avoid:
- excessive horizontal lines,
- full-grid table styling,
- random borders,
- isolated border fragments.

---

## 12. Column widths and row heights

Every key label and value must be readable.

Requirements:
- no important clipped labels,
- no unnecessary wrapping,
- no excessively wide columns,
- no narrow columns that force awkward multi-line headings,
- no row heights that truncate text.

Autofit is a starting point, not a substitute for judgment.

---

## 13. Merged cells

Use merged cells only where they materially improve a section heading.

Avoid merged cells in calculation areas or data tables.

---

## 14. Freeze panes

Use freeze panes when they improve navigation.

Typical:
- freeze top heading rows,
- freeze left label columns on wide schedules.

Do not freeze arbitrary positions.

---

## 15. Zoom and active cell

Set a sensible default zoom.

Typical:
- 90–100% for standard schedules
- 80–90% for wide models

Select a logical starting cell, usually near the upper-left usable portion of the sheet.

---

## 16. Print setup

Where the workbook is meant to be printed or exported:
- define appropriate print areas,
- repeat header rows/columns if needed,
- set orientation appropriately,
- avoid page breaks through key sections,
- check scaling,
- check page titles.

---

## 17. Hidden rows/columns/sheets

Do not hide material content without a clear reason.

If hidden:
- preserve existing intentional controls,
- avoid hiding unresolved issues,
- avoid hiding calculations solely to make a workbook appear simpler.

---

## 18. Charts

Use charts only when they improve interpretation.

Charts should:
- have clear titles,
- use readable labels,
- avoid unnecessary legends,
- avoid 3D effects,
- use restrained formatting.

Do not add charts merely to make a workbook look more polished.

---

## 19. Final XLSX QA

Before delivery:
- workbook opens correctly,
- no formula errors,
- key formulas are intact,
- formulas are consistent,
- no unintended external links,
- inputs are readable,
- output cells are readable,
- tab colors are consistent,
- gridline choice is deliberate,
- widths/heights are appropriate,
- no random borders,
- no clipped text,
- print/layout settings are reasonable,
- final workbook is visually inspected where possible.
