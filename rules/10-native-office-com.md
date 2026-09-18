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
1. Excel COM for advanced/native workbooks and final native operations
2. `openpyxl` for deterministic low-risk structure/formula/style edits
3. direct OOXML for surgical preservation-sensitive edits
4. LibreOffice fallback

### Rendering
1. Office COM native PDF export
2. LibreOffice PDF export

### Validation
1. Python / OOXML structural and formula checks
2. recalculated cached-value checks
3. COM state inspection
4. visual review

## Source safety

- Never open an immutable source for write.
- Never save a source file through Word or Excel.
- Copy the source first.
- All mutating COM operations must target a candidate/working file.
- Disable alerts only for controlled candidate operations.
- Do not suppress a prompt when doing so could accept an unsafe conversion or overwrite.
- Keep link updating disabled by default.
- Do not refresh external queries unless the task explicitly requires it.
- Force-disable Office macros before opening untrusted or task-supplied files.

## COM process ownership

Every automation-created Office instance must be treated as an owned process.

Required lifecycle:
1. create a separate Office instance with `DispatchEx`,
2. capture the process identity immediately from the application window handle,
3. retain an OS process handle for that exact process instance,
4. close owned documents/workbooks in `finally`,
5. call `Quit`,
6. release COM references and run garbage collection,
7. wait briefly for the owned process to exit,
8. if it is still alive, terminate only the retained owned-process handle.

Never:
- use blanket `taskkill /IM EXCEL.EXE`,
- terminate a PID merely because it is named Excel or Word,
- kill an Office process that was not created by the current worker.

The retained process handle is authoritative because it continues to refer to the original process even if a PID is later reused.

## Dialog and timeout safety

`DisplayAlerts = False` does not eliminate all modal states. Activation, Trust Center, Protected View, password, corruption-repair, add-in, printer, and credential dialogs can still block automation.

Therefore:
- every worker call must have a bounded timeout,
- the transport must have a longer outer timeout,
- the Office process must be owned so a timed-out worker can be cleaned up safely,
- timeout errors must be explicit and must not be reported as successful fallback execution.

## Workbook mutation locking

Only one mutating Office operation may target the same canonical workbook path at a time.

Use a cross-process path lock for:
- recalculation + save,
- AutoFit + save,
- Word repagination + save,
- future mutating COM commands.

Read-only rendering/snapshot operations may run concurrently when the underlying file is stable.

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
- circular/iterative calculation validation,
- chart layout,
- conditional formatting/native objects,
- print areas and scaling,
- page setup,
- freeze panes/window state,
- pivot/native workbook features,
- native What-If Data Tables,
- native PDF export.

When calculation logic changes or cached results matter:
1. open the candidate in Excel with links disabled and macros forced off,
2. calculate or `CalculateFullRebuild`,
3. wait until `Application.CalculationState == xlDone`,
4. save,
5. close,
6. reopen/validate with Python using both formula and cached-value views,
7. compare formulas/values/structure/styles against the source where preservation is required.

## AutoFit policy

Do not use blanket AutoFit as final model formatting.

Native AutoFit should be constrained:
- preserve hidden dimensions,
- preserve intentionally narrow spacer columns,
- do not shrink deliberate widths/heights,
- cap only automatic expansion unless the template says otherwise,
- render and visually review important sheets after AutoFit.

## Workbook feature preflight

Before saving an existing workbook through `openpyxl`, inspect the OOXML package for advanced/native features.

If the workbook contains items such as:
- ActiveX,
- embedded objects,
- slicers,
- Power Query/query tables,
- connections,
- complex pivot structures,
- VBA,
- external links,

prefer Excel COM or surgical OOXML. Do not assume a normal `openpyxl` load/save cycle preserves every native Excel object.

## WSL bridge

When the agent runs in WSL:
- convert Linux paths with `wslpath -w`,
- fail closed if path translation fails,
- call the project Windows worker,
- use Windows Python with `pywin32`,
- keep Linux and Windows Python environments separate,
- stage WSL UNC inputs to a per-job local Windows NTFS directory before Office opens them,
- copy successful mutations back atomically,
- stage native render outputs locally when the destination is a WSL UNC path.

Set `OFFICE_COM_PYTHON` if automatic Windows Python discovery fails. Accept either a WSL `/mnt/c/...` path or a standard `C:\...` Windows path and normalize it explicitly.

## Fallback

If COM is unavailable:
- continue with Python/OOXML editing only where feature preflight says it is safe,
- render with LibreOffice,
- report that native Office validation/rendering was unavailable,
- do not pretend fallback rendering is identical to Microsoft Office.
