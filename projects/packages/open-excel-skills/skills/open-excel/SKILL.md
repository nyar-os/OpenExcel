---
name: open-excel
description: Create, revise, review and deliver spreadsheet workbooks with OpenExcel. Use when a task involves workbooks, formulas, pivot summaries, spreadsheet export or workbook review.
---

# open-excel

## Status

Scaffold. The framework is declared and the workbook engine is not implemented.
Every domain operation — opening a workbook, evaluating a formula, aggregating a
pivot, inspecting cells and exporting a format — answers the stable code
`capability/not-implemented`. Do not route around that code with local ZIP, XML,
formula-evaluation or spreadsheet code, and do not present any workbook format
as supported because an extension is recognized.

## Commands

Run `open-excel capabilities` first. `open-excel open <file>` and
`open-excel inspect <file>` currently answer `capability/not-implemented`
because the workbook engine is not wired.

## Preflight

Once the SDK/CLI ships, every task starts here:

1. Query the capabilities the installed SDK/CLI reports and confirm the format,
   sheet, formula, pivot and export entries the task needs. A method name alone
   is not a capability.
2. Confirm the locale and the inputs the task needs (source files, templates,
   data ranges, brand rules) are readable before generating anything.
3. If the preflight reports a missing capability, stop and report it. Never
   degrade silently or swap in an unrelated tool.

## Routes

| Route | Intent | Minimum evidence |
|-------|--------|------------------|
| create | Build a workbook from a brief, data or template | sheet map, formula list, recalculated values, inspect report |
| edit | Change an existing workbook without losing formulas or styles | revision-scoped patch, reopened workbook, inspect diff |
| template | Fill a template through its named ranges or starter sheets | placeholder coverage report, type and range diagnostics |
| review | Audit formulas, references, cycles, types, ranges and styles | inspect report with cell locations, not a single score |
| deliver | Export the target format plus a summary and a loss report | independent reopen, loss report, remaining limitations |

## Hard rules

- Never paste values over formulas to make a report pass.
- Never claim a formula was evaluated or a check ran when it did not.
- Excel or LibreOffice may serve only as an optional independent verification
  source. The workflow must not require either.
- Prefer native formulas, named ranges and pivot definitions over screenshots.
- Report the exact cell address of every finding, not a sheet-level average.

## Safety

- Confirm with the user before any command that writes or overwrites a workbook,
  and state the target path first.
- Never overwrite a source workbook in place unless explicitly asked.
- Never upload workbook contents somewhere the user did not request.

## References

None yet. Design rules, the support matrix and executable examples ship with the
first implemented release.
