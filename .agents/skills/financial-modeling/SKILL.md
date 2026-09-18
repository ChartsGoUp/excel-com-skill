---
name: financial-modeling
description: Apply professional corporate-finance and investment-banking model conventions to XLSX workbooks, including three-statement models, debt schedules, DCF, LBO, valuation, sensitivities, and returns analysis.
---

# Financial Modeling

Read `../../../rules/03-financial-modeling-rules.md` and the XLSX rules.

Enforce:
- clear inputs → schedules → outputs → checks flow,
- explicit historical versus forecast periods,
- transparent assumptions,
- consistent signs and units,
- simple copy-across formulas,
- debt roll-forwards,
- visible model-integrity checks,
- professional number formats,
- restrained use of section colors and input fills.

Do not import assumptions, company facts, dates, or valuation conclusions from the golden example.

Any formula or value change must be intentional and explained. Run the XLSX validator and source comparison before issue when preservation is required.
