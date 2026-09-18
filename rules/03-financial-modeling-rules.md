# 03 — Financial Modeling Rules

## 1. General standard

Use professional investment-banking / corporate-finance model conventions when the workbook is a model or finance schedule.

Priority order:
1. auditability,
2. accounting integrity,
3. logical flow,
4. formula transparency,
5. structural consistency,
6. formatting consistency,
7. visual polish.

A model is not complete because it calculates. It must also be easy to audit and visually consistent.

## 2. Required model architecture

For an integrated three-statement model, use the dependency flow that best fits the business, normally:

**Inputs → Revenue / operating schedules → Working capital → Capex & D&A → Debt & interest → Taxes → Three statements → Valuation / returns → Sensitivities → Checks**

The statements must be outputs of supporting schedules where a schedule exists. Do not independently hardcode or separately forecast statement lines that should link to a driver schedule.

## 3. Workbook and column structure

Define the workbook map before populating formulas.

A financial-model sheet should normally establish:
- title / section identity,
- units,
- row labels,
- historical periods,
- a clear historical/forecast boundary,
- forecast periods,
- optional statistics / CAGR columns,
- controlled spacer columns only where they improve readability.

Do not improvise different period layouts across related schedules without a reason.

Use `config/financial-model-style.json` as the default machine-readable design system unless an approved template overrides it.

## 4. Time periods

Clearly distinguish:
- historical actuals,
- current / interim periods,
- forecast / estimate periods,
- terminal / exit periods.

Period headers must state:
- fiscal year or date,
- actual vs forecast where relevant,
- interim dates where relevant.

Do not label an interim period as a full fiscal year.

Use consistent period granularity across linked schedules.

## 5. Assumptions

Every material assumption should have:
- clear label,
- unit,
- value,
- applicable period,
- source or rationale where useful.

Avoid assumptions embedded invisibly inside formulas.

Centralize assumptions when the same driver affects several schedules. Scenario selectors must be explicit and auditable.

## 6. Formula construction

Prefer:
- simple formulas,
- consistent copy-across logic,
- direct references,
- transparent schedules,
- one economic concept per row where practical.

Avoid:
- unnecessary nested formulas,
- unexplained constants,
- excessive `OFFSET`, `INDIRECT`, or volatile formulas,
- hidden hardcodes,
- formula patterns that change without reason,
- broad `IFERROR(...,0)` when zero is not economically correct.

Equivalent forecast periods should have equivalent normalized R1C1 formula patterns unless explicitly exempted.

## 7. Sign and units

Use one consistent sign convention throughout the model and disclose exceptions.

Typical convention:
- revenue positive,
- expenses negative on statements,
- asset/liability balances positive,
- debt balances positive,
- debt repayments negative cash flow,
- capex negative cash flow.

State units clearly, such as:
- `$`,
- `$000s`,
- `$mm`,
- `%`,
- `x`,
- `days`,
- shares.

Do not mix units inside a schedule without explicit labeling.

## 8. Revenue and operating schedules

Material revenue should be driver-based where source data permits.

Typical drivers include:
- volume × price,
- customers × ARPU,
- locations × sales/location,
- capacity × utilization × price,
- segment growth and mix.

Show material bridges such as:
- volume growth,
- price growth,
- mix,
- margins,
- unit economics.

Revenue in the income statement must link to the revenue schedule.

## 9. Working capital

Where applicable, model material operating working-capital accounts with explicit drivers.

Typical relationships:
- AR from revenue and DSO,
- inventory from COGS and DIO,
- AP from COGS/purchases and DPO,
- other working-capital accounts from the economically relevant base.

State:
- driver,
- day-count convention,
- period timing,
- whether the metric uses average, ending, or another balance basis.

The cash-flow statement change in working capital must reconcile to balance-sheet movements using the selected sign convention.

## 10. PP&E, capex, and depreciation

A standard roll-forward should show:

**Beginning PP&E + Capex + Acquisitions - Disposals - Depreciation ± Other = Ending PP&E**

Where useful, separate:
- maintenance vs growth capex,
- asset classes,
- capex vintages,
- useful lives,
- depreciation methods.

PP&E on the balance sheet and D&A on the income/cash-flow statements must link to the supporting schedule.

## 11. Debt and interest

Debt schedules should clearly show by tranche:
- opening balance,
- borrowings,
- mandatory amortization,
- optional repayment / cash sweep,
- PIK or capitalized interest if applicable,
- other movements,
- closing balance,
- stated/effective rate,
- maturity,
- cash interest,
- fees/OID where relevant.

Opening + movements = closing must be testable.

State the interest base explicitly:
- beginning balance,
- average balance,
- ending balance,
- daily balance,
- another supported convention.

Revolver logic should explicitly incorporate minimum cash and permitted borrowing/repayment behavior.

## 12. Circularity

Do not create accidental circular references.

If circularity is economically required, for example debt ↔ interest ↔ cash sweep:
- use an explicit circularity switch or controlled iterative design,
- document the circular relationship,
- preserve the workbook's intended Excel iteration settings,
- surface iteration settings in native QA,
- provide a non-circular fallback where practical for audit/debugging.

Do not hide circular-reference warnings by hardcoding outputs.

## 13. Taxes

Tax schedules should identify material components where relevant:
- book pre-tax income,
- permanent differences,
- temporary differences,
- NOLs / tax attributes,
- cash taxes,
- deferred taxes,
- jurisdictional or statutory rates where material.

Tax expense, cash taxes, and deferred-tax balance-sheet movements must reconcile to their respective statements.

## 14. Integrated three statements

### Income statement
Statement lines must link to supporting schedules where available.

### Balance sheet
At minimum, ensure:
- cash links to the cash-flow statement,
- working-capital balances link to their schedule,
- PP&E links to the capex/D&A schedule,
- debt links to the debt schedule,
- retained earnings/equity rolls correctly.

Typical retained earnings roll-forward:

**Ending RE = Beginning RE + Net income - Dividends ± Other equity movements**

### Cash-flow statement
At minimum:
- starts from the appropriate earnings measure,
- adds back non-cash items,
- reflects working-capital changes,
- includes investing cash flows,
- includes financing cash flows,
- reconciles beginning to ending cash.

The balance sheet must balance in every modeled period.

## 15. Valuation / returns

Where applicable, show:
- entry assumptions,
- exit assumptions,
- debt,
- cash,
- enterprise value,
- equity value,
- free cash flow,
- WACC,
- terminal assumptions,
- IRR,
- MOIC,
- sensitivity.

Outputs must tie directly to schedules.

## 16. Sensitivities and scenarios

Use sensitivity tables only when decision-useful.

Clearly identify:
- row assumption,
- column assumption,
- base case.

For interactive Excel models, prefer native Excel What-If Data Tables where appropriate.

Scenario systems should use centralized Base / Upside / Downside or equivalent assumptions rather than hardcoded alternative outputs.

## 17. Checks

Every important model must include visible checks.

Examples:
- balance sheet balances,
- cash-flow reconciliation,
- debt roll-forward,
- sources = uses,
- opening + movements = closing,
- retained-earnings roll-forward,
- formula population checks,
- sign consistency,
- historical totals,
- output tie-outs,
- circularity/iteration status where applicable.

Checks should evaluate to zero/OK only when the underlying identity is actually satisfied. Never hardcode a “PASS”.

## 18. Formatting and structure

Formatting is part of model correctness.

Enforce:
- consistent sheet order,
- consistent period columns,
- consistent row hierarchy,
- controlled blank spacing,
- explicit historical/forecast boundary,
- consistent hardcode/formula colors,
- professional number formats,
- restrained borders/fills,
- readable widths/heights,
- deliberate freeze panes,
- deliberate zoom,
- model-appropriate tab colors.

Use semantic styles from `config/financial-model-style.json`. An agent should choose a semantic role such as `section_header`, `hardcode_input`, `subtotal`, or `forecast_header`, not invent visual properties ad hoc.

## 19. QA requirements

Before delivery:
1. run OOXML feature preflight,
2. recalculate with Excel COM when available,
3. wait for calculation completion,
4. validate formulas,
5. validate cached calculated values,
6. compare source/candidate structure and styles when preservation matters,
7. run financial identity checks,
8. render important sheets through native Excel,
9. visually inspect hierarchy, clipping, wrapping, spacing, chart layout, and page setup,
10. correct issues and rerender.

A financial model is not complete until both financial QA and visual QA pass.
