# 03 — Financial Modeling Rules

## 1. General standard

Use professional investment-banking / corporate-finance model conventions when the workbook is a model or finance schedule.

Prioritize:
1. auditability,
2. logical flow,
3. formula transparency,
4. clear assumptions,
5. consistent formatting.

---

## 2. Model flow

Preferred flow:

**Inputs → Operating / supporting schedules → Financial statements / outputs → Checks**

Avoid circular or reverse references unless required and explicitly controlled.

---

## 3. Time periods

Clearly distinguish:
- historical,
- current / interim,
- forecast,
- terminal / exit periods.

Period headers must state:
- fiscal year or period,
- actual vs forecast where relevant,
- interim dates where relevant.

Do not label an interim period as a full fiscal year.

---

## 4. Assumptions

Every material assumption should have:
- clear label,
- unit,
- value,
- period,
- source or rationale where useful.

Avoid assumptions embedded invisibly inside formulas.

---

## 5. Formula construction

Prefer:
- simple formulas,
- consistent copy-across logic,
- direct references,
- transparent schedules.

Avoid:
- unnecessary nested formulas,
- unexplained constants,
- excessive `OFFSET`, `INDIRECT`, or volatile formulas,
- hidden hardcodes,
- formula patterns that change without reason.

---

## 6. Sign convention

Use one consistent sign convention throughout the model.

For example:
- revenue positive,
- expenses negative,
- debt positive balance,
- debt repayment negative cash flow.

Clearly disclose the chosen convention when ambiguity exists.

---

## 7. Historical vs forecast formatting

Where consistent with the selected template:
- historical data should be visually distinguishable from forecast data,
- but use subtle differences rather than heavy colored blocks.

---

## 8. Debt schedules

Debt schedules should clearly show:
- opening balance,
- borrowings,
- repayments,
- interest,
- capitalized interest if applicable,
- cash interest if applicable,
- closing balance,
- interest rate,
- period timing,
- maturity,
- mandatory vs optional repayment if relevant.

Opening + movements = closing must be testable.

---

## 9. Interest

Interest calculations should state:
- rate,
- base,
- timing,
- day-count basis if relevant,
- capitalization mechanics if relevant.

Do not calculate interest from a displayed base that differs from the actual underlying base without disclosure.

---

## 10. Valuation / returns

Where applicable, show:
- entry assumptions,
- exit assumptions,
- debt,
- equity,
- enterprise value,
- cash flow,
- IRR,
- MOIC,
- sensitivity.

Ensure outputs tie directly to schedules.

---

## 11. Checks

Every important model should include checks.

Examples:
- balance sheet balances,
- cash flow reconciliation,
- debt roll-forward,
- sources = uses,
- opening + movements = closing,
- formula population checks,
- sign consistency,
- historical totals,
- output tie-outs.

Checks should fail visibly when something is wrong.

Do not hardcode a “PASS”.

---

## 12. Sensitivities

Use sensitivity tables only when meaningful.

Clearly identify:
- row assumption,
- column assumption,
- base case.

Do not create decorative sensitivity tables.

---

## 13. Delivery standard

A financial model is not complete until:
- formulas calculate,
- key outputs reconcile,
- important assumptions are identifiable,
- checks pass,
- and the workbook has been visually inspected.
