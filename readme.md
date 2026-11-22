# OpenExcel

[![Workbook formats](https://img.shields.io/badge/formats-xlsx%20%7C%20xlsm%20%7C%20csv%20%7C%20tsv-blue)](https://github.com/nyar-os/OpenExcel)
[![License](https://img.shields.io/badge/license-MPL--2.0-blue)](License.md)
[![Node.js](https://img.shields.io/badge/node-%3E%3D18-brightgreen)](projects/packages/open-excel/package.json)

[Package](projects/packages/open-excel/README.md) · [Agent reference](documentation/agents.md)

## 💡 What is OpenExcel?

OpenExcel is a command-line tool and TypeScript SDK for reading, editing, inspecting and exporting spreadsheet workbooks. It keeps cells, formulas, styles, sheet structure and frozen panes in one workbook model, so the same file behaves the same whether it is changed by a person at a terminal, by a script or by an agent.

## 🚀 Getting started

### 1. Agent / prompt (recommended)

Give your coding agent the workbook task in plain language and let it drive the CLI:

```text
Use the `open-excel` CLI to open q1-results.xlsx, check the formulas,
add a Total column with =SUM(B2:D2), sort the rows by Total, and export
a PDF for review.
```

Point the agent at [documentation/agents.md](documentation/agents.md) for the command map, edit-operation JSON shape and stable error codes. Agents should begin every task with `open-excel capabilities --json` so they work with the capabilities of the installed build instead of guessing from file extensions.

### 2. Manual install

Requires Node.js 18 or newer.

```bash
npm install -g @wae/open-excel
open-excel capabilities
```

Create a workbook project, write a formula and export it:

```bash
open-excel create --out ledger.oexcel --title "Ledger"
open-excel edit ledger.oexcel --cell B2 --formula "=A1*2" --out ledger.oexcel --overwrite
open-excel export ledger.oexcel --format xlsx --out ledger.xlsx
```

## ✨ Highlights

- **Cells and ranges.** Read one cell or a rectangle with evaluated values, formulas and cached results; write values, booleans, numbers, errors or formulas back.
- **Formulas and recalculation.** A built-in evaluator covers common aggregation, lookup, text, statistics and math functions. Unknown functions keep the value cached in the source file and are reported as `unsupported` instead of silently changing the result.
- **Formatting and layout.** Apply bold, italic, font size, font color, fill, number format, alignment and wrap to a range; set column widths and row heights; hide rows or columns.
- **Sorting, copying and filling.** Sort whole rows by a column, copy a range to a new anchor, transpose a block, fill a linear series and clear a range.
- **Sheet management.** Add, rename, duplicate, move or remove sheets; switch a sheet between visible, hidden and very hidden; freeze rows and columns or clear the frozen pane.
- **Find.** Search evaluated values and formula sources across a sheet or the whole workbook, with character offsets and an explicit truncation flag.
- **Inspection.** Run rule checks for formulas, references, cycles, types, ranges and styles, and read the findings as data rather than as a pass/fail score only.
- **Export.** Write Excel, CSV, TSV, JSON, HTML and PDF. CSV, TSV, JSON and HTML target the primary sheet and report skipped sheets; PDF is a documented partial projection.
- **Safe edits.** Every edit batch is validated before it is applied, advances the revision once, and can reject a stale writer through an expected revision. Structural edits move merged ranges, frozen panes, row heights, column widths and formula references with the cells.
- **Stable errors.** Failures carry comparable codes such as `document/revision-conflict` and `io/target-exists`, so callers can branch without parsing messages.

## 📊 How OpenExcel compares

| Concern | Ad-hoc spreadsheet scripts | OpenExcel |
| --- | --- | --- |
| Workbook state | Each step reopens and re-exports a file | One workbook model shared by the CLI and SDK |
| Formula results | Stale cached values or an unstated approximation | Recalculated values, with unsupported functions disclosed |
| Overwrites | Silent replacement | Explicit `--overwrite` and an optional revision guard |
| Edits | Hand-built per-task helpers | One JSON edit-operation shape with atomic batches |
| Exports | Separate CSV/HTML/PDF writers | One export surface with loss reports per target |
| Agent use | Screenshots and manual clicking | Machine-readable results and stable error codes |

## 🧪 Examples

| Goal | Entry point |
| --- | --- |
| Read a rectangle with evaluated values | `open-excel range ledger.oexcel A1:D10` |
| Write a formula and save the result | `open-excel edit ledger.oexcel --cell C2 --formula "=A2+B2" --out ledger.oexcel --overwrite` |
| Sort rows by the third column | `open-excel sort ledger.oexcel --range A1:D20 --column 3 --header --out ledger.oexcel --overwrite` |
| Copy a range and keep the source | `open-excel copy ledger.oexcel --range A1:C2 --to E1 --out ledger.oexcel --overwrite` |
| Move a range and clear the source | `open-excel move ledger.oexcel --range A1:C2 --to E1 --out ledger.oexcel --overwrite` |
| Freeze rows and columns at C2 | `open-excel freeze ledger.oexcel --at C2 --out ledger.oexcel --overwrite` |
| Export for review | `open-excel export ledger.oexcel --format pdf --out ledger.pdf` |
| Transpose a block from the SDK | [`range.transpose`](documentation/agents.md#edit-operation-shape) |
| Search values and formulas | `open-excel find ledger.oexcel "invoice"` |

Full command reference: [documentation/agents.md](documentation/agents.md). SDK methods: [@wae/open-excel](projects/packages/open-excel/README.md).

## 📜 License

MPL-2.0. See [License.md](License.md).
