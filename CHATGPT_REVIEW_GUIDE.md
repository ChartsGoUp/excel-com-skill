# Excel COM & Office Artifact Skill — Review Guide for ChatGPT

This repository contains an agent skill and tool system designed to bridge AI coding agents (Claude, Codex, Antigravity, etc.) with **native Microsoft Office COM automation** (specifically Excel and Word) while retaining deterministic Python/OOXML inspection and validation.

---

## 1. What Problems This System Solves

1. **Stale/Missing Formula Values:** Standard Python libraries (`openpyxl`, `xlsxwriter`, `pandas`) can write formula strings (`=SUM(A1:A10)`), but they do not execute Excel's calculation engine. Without native COM, cached values (`<v>` tags in the workbook XML) remain blank or zero until opened by a human.
2. **Column Width Clipping (`###`):** Text metrics cannot be reliably calculated in Python without a full typography rendering engine. COM uses Excel's true Windows font-rendering engine to execute `AutoFit` accurately.
3. **Format Integrity & Institutional Standards:** Prevents agents from generating tacky, high-contrast, non-standard spreadsheets by enforcing investment-banking and accounting rules (Arial 10pt, restrained palette, negatives in parentheses, inputs in blue font / formulas in black font).
4. **Process Leak Prevention:** Headless COM scripts often leave orphaned background Excel processes. This system manages owned process IDs and guarantees clean teardown even upon errors.
5. **WSL-to-Windows Interop:** Enables agents running in Linux/WSL environments to safely trigger native Windows Office COM automation via paths, PowerShell, or dedicated worker bridges.

---

## 2. Repository Structure

| Path | Purpose |
|---|---|
| [`.agents/skills/xlsx-master/SKILL.md`](.agents/skills/xlsx-master/SKILL.md) | Agent skill definition for editing and mastering Excel workbooks using COM and deterministic QA. |
| [`.agents/skills/office-artifact-router/SKILL.md`](.agents/skills/office-artifact-router/SKILL.md) | Universal workspace skill that routes any Office task to the proper engine and rules. |
| [`rules/10-native-office-com.md`](rules/10-native-office-com.md) | Core rules governing native COM execution, process lifecycle, alert suppression, and fallbacks. |
| [`rules/02-xlsx-rules.md`](rules/02-xlsx-rules.md) | Design rules: typography (Arial 10pt), layout, tab colors, formula conventions, number formatting. |
| [`rules/05-qa-checklist.md`](rules/05-qa-checklist.md) | Comprehensive checklist for reviewing artifacts before delivery. |
| [`tools/office_artifacts.py`](tools/office_artifacts.py) | CLI entrypoint for inspection, COM recalculation, autofit, validation, and rendering. |
| [`office_artifact_system/`](office_artifact_system/) | Python package implementing COM bridging (`com_bridge.py`), Excel ops (`xlsx_ops.py`), styling, and QA. |
| [`windows/`](windows/) | Windows COM worker script (`office_worker.py`) and bootstrap automation. |
| [`examples/standalone_excel_com.py`](examples/standalone_excel_com.py) | Self-contained Python context manager demonstrating safe `win32com` headless Excel session management. |

---

## 3. Workflow for Agents

```
1. Source Safety:
   Copy source XLSX to working directory (never edit source in place).

2. Inspect:
   Run `office_artifacts.py inspect-xlsx <path>` to check formulas, sheets, and error states.

3. Edit:
   Use Python/openpyxl to modify structure, formulas, and styles according to institutional rules.

4. Native COM Execution:
   - `com-recalculate-xlsx`: Forces CalculateFullRebuild and updates cached formula results.
   - `com-autofit-xlsx`: Adjusts row and column metrics using Excel's true rendering engine.

5. Validate & Compare:
   - Run `validate-xlsx` to check for #REF!, #VALUE!, or corrupt formulas.
   - Run `compare-xlsx` against the source to ensure no accidental mutations occurred.

6. Visual QA:
   Render to PDF via COM (`render --engine auto`) and visually inspect page layout and overflows.
```

---

## 4. Suggested Areas for Review

If you are reviewing this repository, please evaluate:

1. **Process Safety & Concurrency:** Does the COM session management in `office_artifact_system/com_bridge.py` and `examples/standalone_excel_com.py` adequately handle unexpected crashes, hanging dialogs, or multiple concurrent agents?
2. **WSL-Windows Cross-Platform Bridge:** How robust is the communication between Linux/WSL and Windows Python for COM dispatch? Are there edge cases with path translation (`wslpath`) or quoting?
3. **Institutional Quality & Formatting Standards:** Are the rules in `rules/02-xlsx-rules.md` and `rules/10-native-office-com.md` comprehensive enough for Wall Street / institutional accounting workbooks?
4. **Error Recovery & Fallbacks:** When COM or Excel is unavailable (e.g. running on pure Linux/Docker), how gracefully does the fallback to LibreOffice / Python-only mode behave?
