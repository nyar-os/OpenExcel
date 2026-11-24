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

The range/layout/sheet subcommands remain for compatibility and existing callers. New agents should batch the same changes through the SDK, write numeric values and explicit `{ formula }` cells, recalculate, assert the computed results, inspect the tracked workbook and reopen both the native project and the exported package:

```ts
import assert from 'node:assert/strict';

import { OpenExcel } from '@wae/open-excel';

const excel = new OpenExcel({ locale: 'en' });
const book = await excel.create({ title: 'Ledger' });

// One atomic batch: numeric seeds plus explicit formula cells.
const edited = await excel.edit(book, {
    expectedRevision: book.revision,
    operations: [
        {
            op: 'range.set',
            sheet: 'Sheet1',
            start: 'A1',
            rows: [
                ['Region', 'Q1', 'Q2', 'Total'],
                ['North', 120, 140, null],
                ['South', 90, 110, null],
            ],
        },
        { op: 'cell.set', sheet: 'Sheet1', reference: 'D2', formula: '=B2+C2' },
        { op: 'cell.set', sheet: 'Sheet1', reference: 'D3', formula: '=B3+C3' },
        { op: 'format.set', sheet: 'Sheet1', range: 'A1:D1', style: { bold: true, fill: '#eef2ff' } },
    ],
});

// `edit` evaluates formulas when it commits, so a following `recalc` is
// idempotent and reports `changed: 0`; the cell reads assert the computed results.
const recalc = await excel.recalc(edited.workbook);
assert.equal(recalc.changed, 0);
assert.equal((await excel.cell(edited.workbook, { sheet: 'Sheet1', address: 'D2' })).value, 260);
assert.equal((await excel.cell(edited.workbook, { sheet: 'Sheet1', address: 'D3' })).value, 200);

// Inspect the tracked workbook; a `{ path }` input is also accepted for
// handle-free checks of a package that is never opened.
const inspection = await excel.inspect(edited.workbook, { checks: ['formulas', 'references', 'cycles'] });

// Write the exchange package and the native project, then reopen both.
const xlsx = await excel.export(edited.workbook, { format: 'xlsx', target: 'ledger-final.xlsx', overwrite: true });
await excel.save(edited.workbook, { target: 'ledger-final.oexcel', overwrite: true });
await excel.dispose();

const reopened = new OpenExcel();
for (const path of ['ledger-final.oexcel', 'ledger-final.xlsx']) {
    const restored = await reopened.open({ path });
    const total = await reopened.cell(restored, { sheet: 'Sheet1', address: 'D2' });
    assert.equal(total.formula, '=B2+C2');
    assert.equal(total.value, 260);
}
await reopened.dispose();
console.log(recalc.changed, inspection.passed, xlsx.target);
```

`export` honours `overwrite` like `save` does; without it an existing target is `io/target-exists`. `inspect` reports findings with stable codes and `passed`; treat `error` severity as a failed delivery.

The workflow above exists verbatim as
[`projects/packages/open-excel/examples/ledger-workflow.ts`](../projects/packages/open-excel/examples/ledger-workflow.ts):
it asserts the batch advances the revision once, that the committed `recalc` is
idempotent, that `D2=260` and `D3=200`, and that both the native project and the
exported XLSX keep the formula source and the cached result when reopened. CI
runs it against the real engine via `pnpm --filter @wae/open-excel run example`
after `build:test`.

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
| `workbook.setTitle` | `title` | — | Renames the workbook; the title is trimmed, must not be blank, and shows in summaries and saved projects. |
| `cell.set` | `sheet`, `reference`, and `value` or `formula` | — | Writes one cell. A formula may start with or omit `=`. |
| `cell.clear` | `sheet`, `reference` | `keepStyle` | Clears one cell's value, formula and cached result. With `keepStyle` a styled cell stays with its formatting and an empty value; otherwise the cell is removed entirely. |
| `range.set` | `sheet`, `start`, `rows` | — | Writes a matrix of literals, formulas or `null` to clear. |
| `range.append` | `sheet`, `rows`, `start?` | — | Writes the matrix below the sheet's last populated row; the first column comes from `start` (default `A1`). |
| `range.copy` | `sheet`, `range`, `target` | `targetSheet` | Copies into the rectangle anchored at `target` and keeps the source. `targetSheet` writes into another sheet instead. |
| `range.move` | `sheet`, `range`, `target` | `targetSheet` | Moves into the rectangle anchored at `target` and clears the source. `targetSheet` moves into another sheet and clears the source sheet instead. |
| `range.swap` | `sheet`, `range`, `target` | `targetSheet` | Exchanges the rectangle with the equal-sized rectangle anchored at `target`. `targetSheet` exchanges with another sheet instead. |
| `range.flip` | `sheet`, `range`, `axis` | — | Mirrors the rectangle in place; `horizontal` swaps columns, `vertical` swaps rows. |
| `range.fill` | `sheet`, `range`, `direction` | — | Fills the rectangle from one edge: `down` copies the top row below it, `up` the bottom row above it, `right` the left column rightward, `left` the right column leftward. Formulas copy verbatim without reference adjustment, and merged ranges inside the rectangle are rejected. |
| `range.insert` | `sheet`, `range`, `shift` | — | Inserts an empty rectangle, shifting cells inside its band `down` or `right` by the rectangle's height or width. Moved cells keep styles and formula sources, references elsewhere follow them, merges fully inside the band move with it, and a merge crossing the insertion line is rejected. Heights, widths, freezes and names stay put. |
| `range.delete` | `sheet`, `range`, `shift` | — | Deletes the rectangle, shifting the rest of its band `up` or `left`. References elsewhere follow the moved cells, merges fully inside the band vanish, and a merge crossing the rectangle is rejected. Heights, widths, freezes and names stay put. |
| `range.sort` | `sheet`, `range`, and `column` or `keys` | `ascending`, `hasHeader` | Sorts whole rows by one `column` or an ordered `keys` list; formulas move with their row. |
| `range.removeDuplicates` | `sheet`, `range` | `keys`, `hasHeader` | Keeps the first row per key signature and compacts the survivors upward, clearing the vacated tail. `keys` holds one-based offsets inside the range or column letters and defaults to every column; cells move with their formulas and styles. Appends `{ op, sheet, range, removed }` to `affected`. |
| `range.removeEmptyRows` | `sheet`, `range` | — | Removes rows with no cell anywhere inside the range, compacts the survivors upward and clears the vacated tail. Authored blank cells, empty strings and formulas keep their rows. Appends `{ op, sheet, range, removed }` to `affected`. |
| `range.removeEmptyColumns` | `sheet`, `range` | — | Removes columns with no cell anywhere inside the range, compacts the survivors leftward and clears the vacated right edge. Authored blanks and formulas keep their columns. Appends `{ op, sheet, range, removed }` to `affected`. |
| `range.fillSeries` | `sheet`, `range` | `start`, `step`, `kind`, `ratio` | Fills a single row or column with a linear numeric series, with Excel date serials when `kind` is `date` (the step is then a number of days and each written cell gains the `yyyy-mm-dd` format), or with a geometric series when `kind` is `growth` (each cell multiplies by `ratio`, which defaults to the ratio of the first two seeds, else `2`). |
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
| `name.define` | `name`, `sheet`, `range` | — | Defines a workbook-scope name bound to the rectangle, or redefines the existing name when it matches case-insensitively. The name must not look like a cell reference. |
| `name.remove` | `name` | — | Removes a defined name case-insensitively; an unknown name fails with `object/not-found`. |
| `name.rename` | `name`, `to` | — | Renames a defined name case-insensitively, keeping its sheet and rectangle; an unknown name fails with `object/not-found` and an existing `to` is refused. |
| `sheet.add` | `name` | `index` | Adds a sheet. |
| `sheet.rename` | `sheet`, `name` | — | Renames a sheet. |
| `sheet.remove` | `sheet` | — | Removes a sheet; the last sheet cannot be removed. |
| `sheet.duplicate` | `sheet` | `name` | Duplicates a sheet. |
| `sheet.move` | `sheet`, `index` | — | Moves a sheet to a zero-based position. |
| `row.insert` / `row.delete` | `sheet`, `row` | `count` (default `1`) | Inserts before or deletes from a one-based row. |
| `column.insert` / `column.delete` | `sheet`, `column` | `count` (default `1`) | Inserts before or deletes from a one-based number or letter. |
| `row.swap` / `column.swap` | `sheet`, `row`/`column`, `other` | — | Exchanges two whole rows or columns; cells, heights/widths, single-line merges and single-line named ranges move with the line, formula references to it follow, and a merge straddling either line is rejected. |
| `row.setHeight` | `sheet`, `row`, `height` | `hidden` | Sets row height in points. |
| `column.setWidth` | `sheet`, `column`, `width` | `hidden` | Sets column width in Excel character units. |
| `cell.replaceText` | `sheet`, `find`, `replace` | `range`, `matchCase`, `limit` | Replaces text in string cell values, skipping formulas, and reports the replacement count. `range` defaults to the used rectangle; `limit` caps occurrences in row-major order. Each call appends `{ op: "cell.replaceText", sheet, range, replaced }` to `affected`. |
| `range.setTextCase` | `sheet`, `range`, `case` | — | Rewrites the case of every text cell in the rectangle; `case` is `upper`, `lower` or `title`. Formula cells and non-text values are skipped, and the call appends `{ op: "range.setTextCase", sheet, range, changed }` to `affected`. |
| `range.trimText` | `sheet`, `range` | `collapse` | Trims leading and trailing whitespace from text cells; `collapse: true` also folds internal whitespace runs into single spaces. Formula cells and non-text values are skipped, and the call appends `{ op: "range.trimText", sheet, range, changed }` to `affected`. |
| `range.round` | `sheet`, `range` | `places` (default `0`), `mode` (default `half-up`) | Rounds every numeric value in the rectangle to `places` decimals; `mode` picks the tie direction (`half-up` toward positive infinity, `up` away from zero, `down` toward zero). Formula and text cells are skipped, and the call appends `{ op: "range.round", sheet, range, changed }` to `affected`. |

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
- `range.copy` copies values, formula sources, cached results and styles verbatim; relative references inside formulas are not rewritten. A merged range touching the source or destination is rejected. `targetSheet` selects another destination sheet; the source sheet's cells are not touched, and the target sheet's merges are checked instead of the source rectangle.
- `range.move` snapshots the block, clears both rectangles, then writes the snapshot, so an overlapping destination is deterministic. A merged range touching either rectangle is rejected.
- `range.swap` snapshots both rectangles before clearing either, reports their bounding rectangle as affected on one sheet (the target rectangle when `targetSheet` crosses sheets), and rejects merged ranges touching either side.
- `range.flip` snapshots the rectangle before clearing it, keeps values, formula sources, cached results and styles with their cells, and rejects merged ranges inside the rectangle. A single row flipped vertically or a single column flipped horizontally is a no-op that still reports the rectangle.
- `range.fill` copies the source edge cell by cell, keeping values, formula sources, cached results and styles; formula text is copied as-is, so relative references are not adjusted. A blank source cell clears its target, and the source edge is unchanged. Merged ranges inside the rectangle are rejected. A single row filled up or down and a single column filled left or right are no-ops that still report the rectangle.
- `range.transpose` moves the block, clears the source and rejects a destination that overlaps it or touches a merged range.
- `range.fillSeries` needs exactly one row or one column; merged ranges in the way are rejected. With `kind: "date"` the series writes Excel date serials and stamps `yyyy-mm-dd` on every written cell. `kind: "growth"` multiplies the start by `ratio` per cell; `ratio` defaults to the ratio of the first two numeric seeds, else `2`, must be finite and non-zero, is refused for other kinds, and a series starting at zero needs an explicit ratio.
- `range.sort` sorts numbers before text, text before booleans, and leaves blanks trailing; it refuses ranges that overlap a merged range. `keys` is an ordered list of `{ column, ascending? }` entries compared left to right, and passing both `column` and `keys` is rejected. When `keys` is omitted the single `column` and `ascending` fields describe one key.
- Structural edits (`row.*`, `column.*`) move merged ranges, frozen panes, row heights, column widths and formula references. Deleted reference targets become `#REF!`.
- Defined names (`name.*`) are workbook-scope. They move with row and column inserts and deletes like merged ranges; a name whose whole range is deleted, or whose sheet is removed, is dropped. Names live in the native `.oexcel` project only and are not written to XLSX.
- Unknown formula functions keep the cached source value and set `computedState: "unsupported"`; they never invent a result.
- Legacy and dotted Excel spellings are folded onto engine functions before evaluation: `CONCATENATE`→`CONCAT`, `MODE.SNGL`→`MODE`, `STDEV.S`/`STDEV.P`→`STDEV`/`STDEVP`, `VAR.S`/`VAR.P`→`VAR`/`VARP`, `RANK.EQ`→`RANK`. The stored formula keeps its original text and `capabilities()` lists the folded names beside the engine functions.
- `find` returns at most 500 hits and sets `truncated: true` when more exist.
- `save` and most mutating commands need an explicit `--overwrite` (or `overwrite: true`) to replace an existing file. `export` has no overwrite option at all.
- Close every session with `dispose()`; a disposed session rejects with `session/disposed`.
