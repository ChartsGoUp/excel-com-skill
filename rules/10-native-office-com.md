# 10 — Native Microsoft Office COM Rules

## Purpose

Use native Word and Excel automation when the actual Office rendering/calculation engine materially improves the result.

COM is an execution engine, not a substitute for deterministic validation.

## Engine precedence

### DOCX editing
1. Word COM
2. `python-docx`
3. direct OOXML
4. LibreOffice fallback

### XLSX editing
1. Excel COM
2. `openpyxl`
3. direct OOXML
4. LibreOffice fallback

### Rendering
1. Office COM native PDF export
2. LibreOffice PDF export

### Validation
1. Python / OOXML checks
2. COM state inspection
3. visual review

## Source safety

- Never open an immutable source for write.
- Never save a source file through Word or Excel.
- Copy the source first.
- All mutating COM operations must target a candidate/working file.
- Disable alerts where safe so automation cannot silently block on routine prompts.
- Do not suppress a prompt when doing so could accept an unsafe conversion or overwrite.

## Word COM

Prefer Word COM for:

- table AutoFit,
- page/section layout,
- list formatting,
- style application where Word behavior matters,
- field updates,
- TOC updates,
- repagination,
- header/footer placement,
- native PDF export.

After material layout changes:
1. update fields where appropriate,
2. update TOCs where appropriate,
3. repaginate,
4. save the candidate,
5. export to PDF,
6. inspect the rendered pages.

## Excel COM

Prefer Excel COM for:

- `AutoFit`,
- row heights and column widths,
- workbook calculation,
- `CalculateFullRebuild`,
- chart layout,
- conditional formatting/native objects,
- print areas and scaling,
- page setup,
- freeze panes/window state,
- pivot/native workbook features,
- native PDF export.

When calculation logic changes or cached results matter:
1. open the candidate in Excel,
2. calculate or `CalculateFullRebuild`,
3. save,
4. close,
5. reopen/validate with Python,
6. compare formulas/values against the source where preservation is required.

## Automation hardening

- Use separate Office application instances where practical.
- Set `DisplayAlerts = False` only for controlled candidate operations.
- Ensure every COM application is closed in `finally`.
- Kill no unrelated user Office process.
- Do not use blanket process termination as cleanup.
- Use explicit file paths.
- Keep link updating disabled by default.
- Do not enable macros.
- Do not trust external content.

## WSL bridge

When Codex runs in WSL:
- convert Linux paths with `wslpath -w`,
- call the project Windows worker,
- use Windows Python with `pywin32`,
- keep the Linux and Windows Python environments separate.

Set `OFFICE_COM_PYTHON` if automatic Windows Python discovery fails.

## Fallback

If COM is unavailable:
- continue with Python/OOXML editing where safe,
- render with LibreOffice,
- report that native Office validation/rendering was unavailable,
- do not pretend fallback rendering is identical to Microsoft Office.
