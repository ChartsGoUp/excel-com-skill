# Custom Rules

Use this file for hard user-specific rules that should apply across tasks.

Rules in this file override the generic artifact rules unless the current user request explicitly says otherwise.

---

## Current defaults

### General
- Use a restrained professional finance/accounting aesthetic.
- Avoid formatting that looks decorative, generic, or “AI-generated.”
- Optimize for fast comprehension and auditability.
- Use clear structure instead of excessive colors or boxes.

### DOCX
- Use Arial unless a specific approved template requires another font.
- Use true Word bullets and numbering.
- Make tables easy to read.
- Keep tables on one page where practical, but do not shrink text excessively.
- Avoid unnecessary blank pages and excessive whitespace.
- Do not add a cover page unless requested or justified.
- Preserve source content on formatting-only tasks.

### XLSX
- Use Arial unless a specific approved template requires another font.
- Use professional investment-banking / finance formatting when relevant.
- Use consistent tab colors by role.
- Hide gridlines on polished schedules when appropriate.
- Avoid excessive borders.
- Do not add an executive-summary style section to a calculation schedule unless requested or genuinely useful.
- Ensure input/assumption labels and values are fully readable.
- Do not allow narrow columns to create awkward wrapping or hidden labels.
- Use clean spacing and hierarchy instead of boxing every section.
- Preserve source formulas and values on formatting-only tasks.


---

## Primary golden examples

### DOCX manuals and long-form reports
Use `source-data/golden-examples/Golden - Accounting Manual.docx` as the primary visual quality reference for:
- manuals,
- long accounting reports,
- policy documents,
- section-heavy professional documents.

Prioritize its:
- hierarchy,
- page density,
- formal TOC,
- restrained header/footer,
- table design,
- strong but limited corporate color use.

Do not copy its branding or substantive content.

### Financial-model XLSX
Use `source-data/golden-examples/Golden - Three Statement Financial Model.xlsx` as the primary visual and structural quality
reference for:
- three-statement models,
- DCF,
- LBO,
- debt schedules,
- valuation models,
- investment-banking-style workbooks.

Prioritize its:
- logical sheet flow,
- clean period columns,
- visible assumptions,
- section bars,
- restrained white calculation areas,
- total/subtotal formatting,
- explicit integrity checks.

Do not impose this model's red/yellow visual language on ordinary accounting trackers or registers unless appropriate.
