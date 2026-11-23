# OpenExcel agent reference

How an agent discovers capabilities, drives the `open-excel` CLI and composes the `@wae/open-excel` SDK. The CLI serves short, one-shot tasks (`create`, `open`, `summary`, `cell`, `range`, `inspect`, `export`); multi-step edits and data-driven batches belong in a TypeScript script built on `new OpenExcel()`.

- **GUI** (`open-excel open <file>`, `open-excel gui`): humans inspect data, edit formulas and review formatting interactively.
- **SDK**: batch writes, formula recalculation, inspection loops, rendering and save/reopen.
- **Not delivered yet**: PNG/PDF renderers (PDF export is partial), `xlsm` export, images in PDF, and importing `xls`/`ods`/`xltx`/`xml`.

## Capability discovery

Start every task with the capability matrix, not with a file extension:

```bash
open-excel capabilities
open-excel capabilities --json
```

The matrix reports `engine`, `project`, `formats` (import/edit/export/inspect per extension), `checks`, `formulas.functions`, `formulas.unsupportedBehavior`, `renderers`, `exportTargets` and `exportLimits`. `unsupported` is a truthful answer, not an error. A classified extension can still be unsupported: `ods`, `xls`, `xltx`, `xml` and `pdf` never import in this build.

## CLI entry points

The bin is `open-excel` (`projects/packages/open-excel/dist/cli.js`). A project file is the first positional argument for every command except `capabilities` and `gui`.

| Command | Syntax | Key options |
| --- | --- | --- |
| `capabilities` | `open-excel capabilities` | none |
| `create` | `open-excel create` | `--out <file>` (required), `--title`, `--template` (only `blank`), `--overwrite` |
| `open` | `open-excel open <file>` | none |
| `summary` | `open-excel summary <file>` | none |
| `save` | `open-excel save <file>` | `--out` (defaults to the input), `--overwrite` |
| `edit` | `open-excel edit <file>` | Compatibility only: prefer `excel.edit(book, { operations })`. Options: `--sheet`, `--cell`, `--value`, `--formula`, `--ops`, `--out`, `--overwrite` |
| `sort` | `open-excel sort <file>` | `--sheet`, `--range`, `--column`, `--descending`, `--header`, `--out`, `--overwrite` |
| `copy` | `open-excel copy <file>` | `--sheet`, `--range`, `--to`, `--out`, `--overwrite` |
| `move` | `open-excel move <file>` | `--sheet`, `--range`, `--to`, `--out`, `--overwrite` |
| `swap` | `open-excel swap <file>` | `--sheet`, `--range`, `--to`, `--out`, `--overwrite` |
| `flip` | `open-excel flip <file>` | `--sheet`, `--range`, `--axis <horizontal\|vertical>`, `--out`, `--overwrite` |
| `replace` | `open-excel replace <file>` | `--sheet`, `--range` (defaults to the used rectangle), `--find`, `--replace`, `--match-case`, `--limit`, `--out`, `--overwrite` |
| `format` | `open-excel format <file>` | `--sheet`, `--range`, `--bold`, `--italic`, `--font-size`, `--font-color`, `--fill`, `--number-format`, `--align`, `--wrap`, `--out`, `--overwrite` |
| `layout` | `open-excel layout <file>` | `--sheet`, `--column`, `--width`, `--row`, `--height`, `--hide`, `--out`, `--overwrite` |
| `row` | `open-excel row <file>` | `--sheet`, `--insert <n>` or `--delete <n>`, `--count`, `--out`, `--overwrite` |
| `column` | `open-excel column <file>` | `--sheet`, `--insert <n\|letter>` or `--delete <n\|letter>`, `--count`, `--out`, `--overwrite` |
| `freeze` | `open-excel freeze <file>` | `--sheet`, and exactly one of `--at <cell>`, `--rows <n> --columns <m>`, `--clear`; plus `--out`, `--overwrite` |
| `sheet` | `open-excel sheet <file>` | exactly one of `--add`, `--rename ... --to`, `--remove`, `--duplicate ... --as`, `--move ... --index`, `--hide`, `--show`, `--very-hidden`; plus `--out`, `--overwrite` |
| `inspect` | `open-excel inspect <file>` | `--checks formulas,references,cycles,types,ranges,styles` |
| `cell` | `open-excel cell <file> <A1>` | `--sheet` |
| `range` | `open-excel range <file> <A1:A1>` | `--sheet` |
| `recalc` | `open-excel recalc <file>` | none |
| `find` | `open-excel find <file> <query>` | `--sheet` |
| `export` | `open-excel export <file>` | `--format xlsx\|xlsm\|csv\|tsv\|json\|html\|pdf`, `--out` (required); no `--overwrite` |
| `gui` | `open-excel gui` | none |

Notes:

- Mutating commands (`edit`, `sort`, `copy`, `move`, `swap`, `flip`, `replace`, `format`, `layout`, `row`, `column`, `freeze`, `sheet`) only persist when `--out` is given. Without `--out` the change exists only in the opened handle and is lost when the process exits.
- `--sheet` defaults to `Sheet1`. Addresses are A1-style: `B4`, `A1:C4`.
- `--column` for `sort` is a one-based offset inside `--range`, or an absolute column letter.
- `edit --ops` cannot be combined with `--cell`, `--value` or `--formula`.
- `export` has no overwrite flag; an existing target fails with `io/target-exists`.
- `inspect` exits `1` when a check fails and `0` when everything passes. Usage errors return `2`.

### Global flags

| Flag | Meaning |
| --- | --- |
| `--json` | Write one machine envelope to stdout and keep help/diagnostics on stderr. |
| `--no-interactive` | Disable interactive prompts. |
| `--locale <id>` | Select a locale from `locales/locales.json` (only `en-US` ships today). |
| `--help`, `-h` | Print derived help for the current command. |

JSON mode uses the `wae.cli-result/v0` envelope:

```json
{ "schema": "wae.cli-result/v0", "ok": true, "data": { } }
{ "schema": "wae.cli-result/v0", "ok": false, "error": { "code": "document/revision-conflict", "message": "..." } }
```

## Recommended workflows

A short user-facing session starts with one command:

```bash
open-excel create --out ledger.oexcel --title Ledger --overwrite
open-excel open ledger.oexcel
```

The range/layout/sheet subcommands remain for compatibility and existing callers. New agents should batch the same changes through the SDK, recalculate explicitly, inspect the result and reopen the saved project:

```ts
import { OpenExcel } from '@wae/open-excel';

const excel = new OpenExcel({ locale: 'en' });
const book = await excel.open({ path: 'ledger.xlsx' });

// One atomic batch: seed rows, add formulas, style the header.
await excel.edit(book, {
    expectedRevision: book.revision,
    operations: [
        {
            op: 'range.set',
            sheet: 'Sheet1',
            start: 'A1',
            rows: [
                ['Region', 'Q1', 'Q2', 'Total'],
                ['North', '120', '140', '=B2+C2'],
                ['South', '90', '110', '=B3+C3'],
            ],
        },
        { op: 'format.set', sheet: 'Sheet1', range: 'A1:D1', style: { bold: true, fill: '#eef2ff' } },
    ],
});

const recalc = await excel.recalc(book);
const inspection = await excel.inspect(book, { checks: ['formulas', 'references', 'cycles'] });

const xlsx = await excel.export(book, { format: 'xlsx', target: 'ledger-final.xlsx', overwrite: true });
await excel.save(book, { target: 'ledger-final.oexcel', overwrite: true });
await excel.dispose();

// Reopen the native project instead of trusting the writer.
const reopened = new OpenExcel();
const again = await reopened.open({ path: 'ledger-final.oexcel' });
console.log(recalc.changed, inspection.passed, xlsx.target, again.revision);
await reopened.dispose();
```

`export` refuses to overwrite: delete the old file or choose a new target. `inspect` reports findings with stable codes and `passed`; treat `error` severity as a failed delivery.

## Edit operation shape

`Workbook.edit()` and `edit --ops` take:

```json
{
  "expectedRevision": 3,
  "operations": [
    { "op": "cell.set", "sheet": "Sheet1", "reference": "A1", "value": 42 },
    { "op": "cell.set", "sheet": "Sheet1", "reference": "B1", "formula": "=A1*2" }
  ]
}
```

`expectedRevision` is optional. When present and stale, the whole batch fails with `document/revision-conflict` and the workbook is unchanged.

| `op` | Required fields | Optional fields | Effect |
| --- | --- | --- | --- |
| `cell.set` | `sheet`, `reference`, and `value` or `formula` | — | Writes one cell. A formula may start with or omit `=`. |
| `range.set` | `sheet`, `start`, `rows` | — | Writes a matrix of literals, formulas or `null` to clear. |
| `range.append` | `sheet`, `rows`, `start?` | — | Writes the matrix below the sheet's last populated row; the first column comes from `start` (default `A1`). |
| `range.copy` | `sheet`, `range`, `target` | — | Copies into the rectangle anchored at `target` and keeps the source. |
| `range.move` | `sheet`, `range`, `target` | — | Moves into the rectangle anchored at `target` and clears the source. |
| `range.swap` | `sheet`, `range`, `target` | — | Exchanges the rectangle with the equal-sized rectangle anchored at `target`. |
| `range.flip` | `sheet`, `range`, `axis` | — | Mirrors the rectangle in place; `horizontal` swaps columns, `vertical` swaps rows. |
| `range.sort` | `sheet`, `range`, `column` | `ascending`, `hasHeader` | Sorts whole rows; formulas move with their row. |
| `range.fillSeries` | `sheet`, `range` | `start`, `step` | Fills a single row or column with a linear numeric series. |
| `range.transpose` | `sheet`, `range`, `target` | — | Moves the block into the transposed rectangle and clears the source. |
| `range.clear` | `sheet`, `range` | — | Removes every cell in the rectangle, including styles and formulas. |
| `range.clearContents` | `sheet`, `range` | — | Removes values and formulas in the rectangle, keeping every cell style. |
| `format.copy` | `sheet`, `source` (one cell), `target` | Copies the source cell's style onto every target cell, keeping values and formulas. |
| `format.set` | `sheet`, `range`, `style` | — | Applies `bold`, `italic`, `fontSize`, `fontColor`, `fill`, `numberFormat`, `align`, `wrap`. |
| `format.clear` | `sheet`, `range` | — | Removes every cell style in the rectangle, keeping values and formulas. |
| `format.number` | `sheet`, `range`, `kind` | `decimals`, `symbol` | Applies the number format built from `kind`: `currency` (`$#,##0.00`, `symbol` replaces `$`), `percent` (`0.00%`), `date` (`yyyy-mm-dd`) or `plain` (`General`); `decimals` (0-12) sets the fraction digits. Other style properties stay. |
| `merge.set` | `sheet`, `start`, `end`, `merge` | — | Merges or unmerges the rectangle. |
| `sheet.freeze` | `sheet`, and `cell` or `rows`+`columns` or `clear` | — | Freezes rows/columns or clears the pane. `C2` means 1 row and 2 columns. |
| `sheet.setVisibility` | `sheet`, `visibility` | — | `visible`, `hidden` or `veryHidden`; the last visible sheet cannot be hidden. |
| `sheet.setAutoFilter` | `sheet`, `range` | — | Declares the sheet autofilter (an A1 range such as `A1:D20`); `null` clears it. |
| `sheet.add` | `name` | `index` | Adds a sheet. |
| `sheet.rename` | `sheet`, `name` | — | Renames a sheet. |
| `sheet.remove` | `sheet` | — | Removes a sheet; the last sheet cannot be removed. |
| `sheet.duplicate` | `sheet` | `name` | Duplicates a sheet. |
| `sheet.move` | `sheet`, `index` | — | Moves a sheet to a zero-based position. |
| `row.insert` / `row.delete` | `sheet`, `row` | `count` (default `1`) | Inserts before or deletes from a one-based row. |
| `column.insert` / `column.delete` | `sheet`, `column` | `count` (default `1`) | Inserts before or deletes from a one-based number or letter. |
| `row.setHeight` | `sheet`, `row`, `height` | `hidden` | Sets row height in points. |
| `column.setWidth` | `sheet`, `column`, `width` | `hidden` | Sets column width in Excel character units. |
| `cell.replaceText` | `sheet`, `find`, `replace` | `range`, `matchCase`, `limit` | Replaces text in string cell values, skipping formulas, and reports the replacement count. `range` defaults to the used rectangle; `limit` caps occurrences in row-major order. Each call appends `{ op: "cell.replaceText", sheet, range, replaced }` to `affected`. |

The batch is atomic: every operation is validated before the first one is applied, and a successful batch advances `revision` once. `sheet` accepts a sheet name or id.

## Stable error codes

`OpenExcelError.code` is one of:

```text
session/disposed
session/argument-invalid
session/budget-exceeded
workbook/not-found
document/revision-conflict
io/target-exists
io/failure
object/not-found
object/unsupported
format/unsupported
format/invalid
capability/not-implemented
```

CLI usage failures add `open-excel.cli.*` codes such as `open-excel.cli.file-required` and `open-excel.cli.out-required`. Framework failures use `commander.*`. Branch on `code`, never on the message text.

## Key types and files

| Path | Contents |
| --- | --- |
| `projects/packages/open-excel/src/index.ts` | Public exports: `OpenExcel`, errors, format registry, GUI helpers and all types. |
| `projects/packages/open-excel/src/types.ts` | `Workbook`, `EditOperation`, `EditRequest`, `CellStyleInput`, `ExportFormat`, `FindResult`, `SupportLevel`, capability matrix. |
| `projects/packages/open-excel/src/errors.ts` | `errorCodes`, `OpenExcelError`, `NotImplementedError`. |
| `projects/packages/open-excel/src/formats.ts` | `workbookFormats` and the classification registry. Classification is not a support claim. |
| `projects/packages/open-excel/src/open-excel.ts` | The session: `capabilities`, `create`, `open`, `save`, `summary`, `edit`, `cell`, `range`, `recalc`, `find`, `export`, `render`, `dispose`. |
| `projects/packages/open-excel/src/service.ts` | One function per native call. |
| `projects/packages/open-excel/src/cli.ts` | The command tree. Help text lives in `locales/en-US/cli.json`. |
| `projects/crates/model/src/lib.rs` | `capabilities`, `create`, `open`, `save`, `inspect`, `export`, `cell`, `range`, `find`, `recalculate`, `edit`, `render`, `close`. |
| `projects/crates/model/src/eval.rs` | Formula evaluation and cycle detection. |
| `projects/crates/model/src/xlsx.rs`, `csv.rs`, `pdf.rs`, `render.rs` | Importers, writers and projections. |
| `projects/crates/model/src/shift.rs` | Row/column structural shifting and formula-reference rewriting. |
| `projects/crates/napi/src/lib.rs` | JSON-in/JSON-out Node-API exports. |

## Guardrails

- Read `capabilities()` before promising a format, check, renderer or export target. Never infer support from an extension or a method name.
- `open()` accepts a path only. In-memory `bytes` and intents other than `edit` are rejected with `format/unsupported`.
- Supported imports: `.oexcel`, `.xlsx`, `.xlsm`, `.csv`, `.tsv`. `.xlsm` imports but does not export.
- Export targets: `xlsx`, `csv`, `tsv`, `json`, `html`, `pdf`. CSV, TSV, JSON and HTML write the primary sheet only and report every other sheet as a loss. PDF is partial (`style-approximated`, `no-images`).
- `range.copy` copies values, formula sources, cached results and styles verbatim; relative references inside formulas are not rewritten. A merged range touching the source or destination is rejected.
- `range.move` snapshots the block, clears both rectangles, then writes the snapshot, so an overlapping destination is deterministic. A merged range touching either rectangle is rejected.
- `range.swap` snapshots both rectangles before clearing either, reports their bounding rectangle as affected, and rejects merged ranges touching either side.
- `range.flip` snapshots the rectangle before clearing it, keeps values, formula sources, cached results and styles with their cells, and rejects merged ranges inside the rectangle. A single row flipped vertically or a single column flipped horizontally is a no-op that still reports the rectangle.
- `range.transpose` moves the block, clears the source and rejects a destination that overlaps it or touches a merged range.
- `range.fillSeries` needs exactly one row or one column; merged ranges in the way are rejected.
- `range.sort` sorts numbers before text, text before booleans, and leaves blanks trailing; it refuses ranges that overlap a merged range.
- Structural edits (`row.*`, `column.*`) move merged ranges, frozen panes, row heights, column widths and formula references. Deleted reference targets become `#REF!`.
- Unknown formula functions keep the cached source value and set `computedState: "unsupported"`; they never invent a result.
- `find` returns at most 500 hits and sets `truncated: true` when more exist.
- `save` and most mutating commands need an explicit `--overwrite` (or `overwrite: true`) to replace an existing file. `export` has no overwrite option at all.
- Close every session with `dispose()`; a disposed session rejects with `session/disposed`.
