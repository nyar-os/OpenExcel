/**
 * Workbook types of `@wae/open-excel`.
 *
 * This is the frozen P0 shape of the surface: `new OpenExcel(options?)` creates
 * a session, `create`/`open` return a `Workbook` summary, and every sheet keeps
 * a stable identity plus an index used only for ordering.
 *
 * Addresses are A1-style strings. Geometry and pixels never enter the model; a
 * future renderer projects the engine's own layout.
 */

import type { WorkbookFormat, WorkbookIntent } from './formats.js';

/** Opaque workbook identity. */
export type WorkbookId = string;
/** Opaque sheet identity; the index is order only. */
export type SheetId = string;
/** A1-style cell address such as `B4`. */
export type CellAddress = string;

/** Hard limits applied before work starts. */
export interface ResourceBudgets {
    readonly maxWorkbookBytes?: number;
    readonly maxSheetCount?: number;
    readonly maxCellCount?: number;
}

/** Structured diagnostic emitted during open, inspection or export. */
export interface WorkbookDiagnostic {
    readonly code: string;
    readonly severity: 'info' | 'warning' | 'error';
    readonly revision?: number;
    readonly sheetId?: SheetId;
    readonly cell?: CellAddress;
    readonly message: string;
}

/** Session options. Constructing a session opens no window, file or socket. */
export interface OpenExcelOptions {
    /** BCP-47 tag used for diagnostics and default copy. Defaults to `en`. */
    readonly locale?: string;
    readonly budgets?: ResourceBudgets;
    readonly diagnostics?: (event: WorkbookDiagnostic) => void;
}

/** Disclosure level. `partial` always ships a loss path and diagnostic codes. */
export type SupportLevel = 'supported' | 'partial' | 'unsupported';

/** Inspection rules understood by `workbook.inspect()`. */
export type InspectCheck = 'formulas' | 'references' | 'cycles' | 'types' | 'ranges' | 'styles';

/** Per-sheet render targets. */
export type RenderFormat = 'svg' | 'png' | 'pdf';

/** Formats `Workbook.export()` can write. */
export type ExportFormat = 'xlsx' | 'xlsm' | 'csv' | 'tsv' | 'json' | 'html' | 'pdf';

/** What the installed engine can do per format. */
export interface FormatCapability {
    readonly import: SupportLevel;
    readonly edit: SupportLevel;
    readonly export: SupportLevel;
    readonly inspect: SupportLevel;
    /** Codes reported when a `partial` path loses something. */
    readonly lossCodes: readonly string[];
}

/** Result of `OpenExcel.capabilities()`. Read this instead of assuming a method works. */
export interface OpenExcelCapabilities {
    readonly engine: { readonly package: string; readonly version: string };
    /** Whole-project container support; `unsupported` while no engine ships. */
    readonly project: SupportLevel;
    readonly formats: Readonly<Record<WorkbookFormat, FormatCapability>>;
    readonly checks: Readonly<Record<InspectCheck, SupportLevel>>;
    /** Formula functions the engine evaluates and how it treats the rest. */
    readonly formulas: FormulaCapability;
    readonly renderers: Readonly<Record<RenderFormat, SupportLevel>>;
    readonly exportTargets: Readonly<Record<ExportFormat, SupportLevel>>;
    /** What each `partial` export target approximates or drops. */
    readonly exportLimits?: Readonly<Partial<Record<ExportFormat, readonly string[]>>>;
}

/** Formula disclosure of the installed engine. */
export interface FormulaCapability {
    /** Function names the evaluator implements, uppercase. */
    readonly functions: readonly string[];
    /**
     * `preserve-cached` means a formula using an unknown function keeps the
     * value the source workbook cached instead of `#NAME?`.
     */
    readonly unsupportedBehavior: 'preserve-cached';
}

/** Options of `OpenExcel.create()`. */
export interface WorkbookCreateOptions {
    readonly title?: string;
    /** Template name or path the engine should start from. */
    readonly template?: string;
    readonly locale?: string;
}

/**
 * Input of `OpenExcel.open()`.
 *
 * Exactly one of `path` or `bytes` must be present, so the engine never guesses
 * whether a string is a file name or serialized content.
 */
export type WorkbookInput =
    | { readonly path: string; readonly intent?: WorkbookIntent }
    | { readonly bytes: Uint8Array; readonly filename?: string; readonly intent?: WorkbookIntent };

/** One sheet of a workbook summary. */
export interface SheetSummary {
    readonly id: SheetId;
    readonly index: number;
    readonly name: string;
    readonly hidden: boolean;
    /** Non-empty cells on this sheet. */
    readonly cells: number;
}

/** One cell, used when a caller asks for a single address. */
export interface CellSummary {
    readonly sheetId: SheetId;
    readonly address: CellAddress;
    readonly kind: 'text' | 'number' | 'boolean' | 'error' | 'empty' | 'formula';
    readonly formula: string | null;
    readonly value: string | number | boolean | null;
}

/**
 * Opened workbook summary. `revision` is monotonic and changes on every
 * committed mutation, so callers can reject a stale writer.
 */
export interface Workbook {
    readonly id: WorkbookId;
    readonly revision: number;
    readonly title: string | null;
    readonly sheets: readonly SheetSummary[];
    /** Total non-empty cells across every sheet. */
    readonly cells: number;
}

/** Options of `OpenExcel.inspect()`. */
export interface InspectOptions {
    readonly checks?: readonly InspectCheck[];
}

/** One finding of an inspection run. */
export interface InspectIssue {
    readonly code: string;
    readonly severity: 'info' | 'warning' | 'error';
    readonly revision: number;
    readonly sheetId?: SheetId;
    readonly cell?: CellAddress;
    readonly recoverable: boolean;
    readonly message: string;
}

/** Result of `OpenExcel.inspect()` or `Workbook.inspect()`. Rules only, never a score. */
export interface InspectReport {
    readonly revision: number;
    readonly checkerVersion: string;
    readonly issues: readonly InspectIssue[];
    readonly passed: boolean;
}

/** Options of saving a native project (not an exchange format). */
export interface SaveOptions {
    readonly target: string;
    readonly overwrite?: boolean;
}

/** Result of a save. */
export interface SaveReport {
    readonly target: string;
    readonly revision: number;
    readonly sheets: number;
    readonly bytes: number;
}

/** Options of `Workbook.export()`. */
export interface ExportOptions {
    readonly format: ExportFormat;
    readonly target: string;
    readonly overwrite?: boolean;
}

/** One thing a projection could not keep. */
export interface ExportLoss {
    readonly code: string;
    readonly severity: 'warning' | 'error';
    readonly sheetId?: SheetId;
    readonly cell?: CellAddress;
    readonly message: string;
}

/** Result of a projection. Reports what the target actually kept. */
export interface ExportReport {
    readonly revision: number;
    readonly format: ExportFormat;
    readonly target: string;
    readonly losses: readonly ExportLoss[];
    /** False for projections such as PDF that cannot be reopened for editing. */
    readonly editable: boolean;
}

/** Options of rendering one or more sheets. */
export interface RenderOptions {
    readonly format: RenderFormat;
    readonly sheets?: 'all' | readonly SheetId[];
    readonly target?: string;
}

/** One rendered sheet. */
export interface RenderPage {
    readonly sheetId: SheetId;
    readonly path?: string;
    readonly bytes?: number;
}

/** Result of a render run. */
export interface RenderReport {
    readonly revision: number;
    readonly format: RenderFormat;
    readonly pages: readonly RenderPage[];
    readonly diagnostics: readonly WorkbookDiagnostic[];
}

/** Options of `Workbook.render()`. */
export interface WorkbookRenderOptions {
    readonly format?: RenderFormat;
    /** Sheet id or name; the first visible sheet when omitted. */
    readonly sheet?: string;
}

/** SVG result of rendering one sheet. */
export interface SvgRender {
    readonly revision: number;
    readonly format: 'svg';
    readonly sheet: string;
    readonly sheetId: SheetId;
    readonly svg: string;
}

/** Value accepted by `cell.set`. */
export type CellValueInput =
    | string
    | number
    | boolean
    | null
    | { readonly kind: 'text'; readonly text: string }
    | { readonly kind: 'number'; readonly number: number }
    | { readonly kind: 'bool'; readonly value: boolean }
    | { readonly kind: 'error'; readonly code: string }
    | { readonly kind: 'empty' };

/** JSON projection of one evaluated cell value. Errors keep their literal. */
export type CellPrimitive = string | number | boolean | null;

/** What one cell holds after evaluation. */
export type CellKind = 'text' | 'number' | 'boolean' | 'error' | 'empty' | 'formula';

/** Result of reading one cell. `value` is the evaluated result. */
export interface CellDetail {
    readonly sheet: string;
    readonly address: CellAddress;
    readonly kind: CellKind;
    readonly formula: string | null;
    /** Evaluated value: what a reader shows. */
    readonly value: CellPrimitive;
    /** Value the source workbook carried next to the formula, when it had one. */
    readonly cached: CellPrimitive;
    /**
     * How the last evaluation ended: `unsupported` when the engine kept the
     * cached value because it cannot evaluate a function in the formula.
     */
    readonly computedState: 'computed' | 'unsupported' | null;
}

/** One cell of a range matrix. */
export interface RangeCell {
    readonly address: CellAddress;
    readonly value: CellPrimitive;
    readonly formula: string | null;
}

/** Result of reading a rectangle, row by row. */
export interface RangeMatrix {
    readonly sheet: string;
    readonly range: string;
    readonly rows: readonly (readonly RangeCell[])[];
}

/** Result of `Workbook.recalc()`. */
export interface RecalcReport {
    readonly revision: number;
    /** Formula cells whose evaluated value changed. */
    readonly changed: number;
}

/** Options of reading one cell. */
export interface CellQuery {
    /** Sheet name or id; the first visible sheet when omitted. */
    readonly sheet?: string;
    readonly address: CellAddress;
}

/** Options of reading one rectangle. */
export interface RangeQuery {
    /** Sheet name or id; the first visible sheet when omitted. */
    readonly sheet?: string;
    /** A1 range such as `A1:C4`. */
    readonly range: string;
}

/** One `range.set` entry: a literal, `null` to clear, or a formula with an optional cached value. */
export type CellInput = CellValueInput | { readonly formula: string; readonly value?: CellValueInput };

/** Visual properties accepted by `format.set`. Omitted properties keep their current value. */
export interface CellStyleInput {
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly fontSize?: number;
    readonly fontColor?: string;
    readonly fill?: string;
    readonly numberFormat?: string;
    readonly align?: 'left' | 'center' | 'right' | 'justify' | 'distribute' | 'fill';
    readonly wrap?: boolean;
}

/** Sheet tab visibility accepted by `sheet.setVisibility`. */
export type SheetVisibility = 'visible' | 'hidden' | 'veryHidden';

/** Parsed options of the `freeze` command; exactly one mode applies. */
export interface FreezeCommandOptions {
    readonly sheet?: string;
    /** A1 cell that marks the freeze boundary, such as `C2`. */
    readonly at?: CellAddress;
    /** Number of frozen rows when no `--at` cell is given. */
    readonly rows?: number;
    /** Number of frozen columns when no `--at` cell is given. */
    readonly columns?: number;
    /** Clear the frozen pane. */
    readonly clear?: boolean;
}

/** Parsed options of the `sheet` visibility flags; exactly one name applies. */
export interface SheetVisibilityCommandOptions {
    readonly hide?: string;
    readonly show?: string;
    readonly veryHidden?: string;
}

/** One rectangle an edit operation rewrote, as reported by `Workbook.edit()`. */
export interface AffectedRange {
    readonly op:
        | 'range.sort'
        | 'range.fillSeries'
        | 'range.transpose'
        | 'range.copy'
        | 'range.move'
        | 'range.swap'
        | 'range.flip'
        | 'cell.replaceText';
    readonly sheet: string;
    /** A1 rectangle such as `A1:A5`. */
    readonly range: string;
    /** Occurrences replaced by a `cell.replaceText` operation. */
    readonly replaced?: number;
}

/**
 * Copies `range` into the rectangle anchored at `target` and keeps the source.
 * Values, formula sources, cached results and styles move verbatim, while
 * relative references inside formulas are not rewritten. The source is
 * snapshotted before writing, so overlapping source and target stay
 * deterministic; merged ranges intersecting either rectangle are rejected.
 */
export interface RangeCopyOperation {
    readonly op: 'range.copy';
    readonly sheet: string;
    readonly range: string;
    /** Top-left cell of the destination rectangle. */
    readonly target: CellAddress;
}

/**
 * Moves `range` into the rectangle anchored at `target` and clears the source.
 * Values, formula sources, cached results and styles move verbatim, while
 * relative references inside formulas are not rewritten. The source is
 * snapshotted before either rectangle is cleared, so overlapping source and
 * target stay deterministic; merged ranges intersecting either rectangle are
 * rejected.
 */
export interface RangeMoveOperation {
    readonly op: 'range.move';
    readonly sheet: string;
    readonly range: string;
    /** Top-left cell of the destination rectangle. */
    readonly target: CellAddress;
}

/**
 * Exchanges `range` with the equal-sized rectangle anchored at `target`.
 * Values, formula sources, cached results and styles travel verbatim, while
 * relative references inside formulas are not rewritten. Both rectangles are
 * snapshotted before either is cleared, so overlapping rectangles stay
 * deterministic; merged ranges intersecting either rectangle are rejected.
 */
export interface RangeSwapOperation {
    readonly op: 'range.swap';
    readonly sheet: string;
    readonly range: string;
    /** Top-left cell of the second, equal-sized rectangle. */
    readonly target: CellAddress;
}

/**
 * Mirrors `range` in place along `axis`.
 *
 * `horizontal` swaps columns left to right while rows keep their position;
 * `vertical` swaps rows top to bottom while columns keep theirs. Cell contents,
 * styles and formula sources travel verbatim, while relative references inside
 * formulas are not rewritten. The rectangle is snapshotted before it is
 * cleared; merged cells inside the rectangle are rejected.
 */
export interface RangeFlipOperation {
    readonly op: 'range.flip';
    readonly sheet: string;
    readonly range: string;
    /** `horizontal` mirrors columns; `vertical` mirrors rows. */
    readonly axis: 'horizontal' | 'vertical';
}

/** One operation of `Workbook.edit()`. */
export type EditOperation =
    | {
          readonly op: 'cell.set';
          readonly sheet: string;
          readonly reference: CellAddress;
          readonly value?: CellValueInput;
          readonly formula?: string;
      }
    | {
          readonly op: 'range.set';
          readonly sheet: string;
          /** Top-left cell of the matrix. */
          readonly start: CellAddress;
          readonly rows: readonly (readonly CellInput[])[];
      }
    | {
          /**
           * Appends rows below the sheet's last populated row, starting at
           * `start` (default `A1`). The first column is `start`'s column.
           */
          readonly op: 'range.append';
          readonly sheet: string;
          /** Top-left cell for the appended block; defaults to `A1`. */
          readonly start?: CellAddress;
          readonly rows: readonly (readonly CellInput[])[];
      }
    | { readonly op: 'format.set'; readonly sheet: string; readonly range: string; readonly style: CellStyleInput }
    | {
          /** Removes values and formulas in `range` while every cell style stays in place. */
          readonly op: 'range.clearContents';
          readonly sheet: string;
          /** A1 range such as `A1:C4`. */
          readonly range: string;
      }
    | {
          /** Removes every cell style in `range`; values and formulas stay in place. */
          readonly op: 'format.clear';
          readonly sheet: string;
          /** A1 range such as `A1:C4`. */
          readonly range: string;
      }
    | {
          /** Copies the style of one cell onto every cell in `target`, values untouched. */
          readonly op: 'format.copy';
          readonly sheet: string;
          /** One-cell source such as `B2`; it must already carry content or a style. */
          readonly source: string;
          /** A1 target range such as `B2:D8`. */
          readonly target: string;
      }
    | {
          /** Merges `start:end` when `merge` is true and unmerges it otherwise. */
          readonly op: 'merge.set';
          readonly sheet: string;
          readonly start: CellAddress;
          readonly end: CellAddress;
          readonly merge: boolean;
      }
    | {
          /**
           * Freezes rows and columns of one sheet. A `cell` boundary such as `C2`
           * freezes the rows above it and the columns left of it, so `C2` means
           * `rows: 1` and `columns: 2`. Explicit `rows` and `columns` are counts.
           * `clear: true`, or omitting every boundary, removes the frozen pane.
           */
          readonly op: 'sheet.freeze';
          readonly sheet: string;
          /** A1 cell whose top-left corner is the freeze boundary. */
          readonly cell?: CellAddress;
          /** Number of frozen rows when no `cell` is given. */
          readonly rows?: number;
          /** Number of frozen columns when no `cell` is given. */
          readonly columns?: number;
          /** Removes the frozen pane. */
          readonly clear?: boolean;
      }
    | {
          /** Sets one column width in Excel character units; `hidden` hides the column. */
          readonly op: 'column.setWidth';
          readonly sheet: string;
          /** One-based column number or a column letter such as `B`. */
          readonly column: number | string;
          readonly width: number;
          readonly hidden?: boolean;
      }
    | {
          /** Sets one row height in points; `hidden` hides the row. */
          readonly op: 'row.setHeight';
          readonly sheet: string;
          /** One-based row number. */
          readonly row: number;
          readonly height: number;
          readonly hidden?: boolean;
      }
    | { readonly op: 'sheet.setVisibility'; readonly sheet: string; readonly visibility: SheetVisibility }
    | {
          /**
           * Inserts `count` empty rows before `row`, moving cells, merged ranges,
           * frozen panes, row heights, column widths and local formula references.
           */
          readonly op: 'row.insert';
          readonly sheet: string;
          /** One-based row the new rows appear before. */
          readonly row: number;
          /** Number of rows to insert; defaults to 1. */
          readonly count?: number;
      }
    | {
          /** Deletes `count` rows starting at `row`; the band must stay inside the sheet. */
          readonly op: 'row.delete';
          readonly sheet: string;
          /** One-based first row to delete. */
          readonly row: number;
          /** Number of rows to delete; defaults to 1. */
          readonly count?: number;
      }
    | {
          /** Inserts `count` empty columns before `column`, moving every column-bearing field. */
          readonly op: 'column.insert';
          readonly sheet: string;
          /** One-based column number or a column letter such as `B`. */
          readonly column: number | string;
          /** Number of columns to insert; defaults to 1. */
          readonly count?: number;
      }
    | {
          /** Deletes `count` columns starting at `column`; the band must stay inside the sheet. */
          readonly op: 'column.delete';
          readonly sheet: string;
          /** One-based column number or a column letter such as `B`. */
          readonly column: number | string;
          /** Number of columns to delete; defaults to 1. */
          readonly count?: number;
      }
    | {
          /**
           * Replaces text inside `range` and reports the replacement count.
           * When `range` is omitted the sheet's used rectangle is scanned.
           * Formula cells are skipped; `limit` caps the occurrences replaced,
           * scanning cells in row-major order.
           */
          readonly op: 'cell.replaceText';
          readonly sheet: string;
          /** A1 rectangle such as `A1:C4`; defaults to the used rectangle. */
          readonly range?: string;
          readonly find: string;
          readonly replace: string;
          readonly matchCase?: boolean;
          /** Maximum occurrences to replace; unlimited when omitted. */
          readonly limit?: number;
      }
    | { readonly op: 'sheet.add'; readonly name: string; readonly index?: number }
    | { readonly op: 'sheet.rename'; readonly sheet: string; readonly name: string }
    | { readonly op: 'sheet.remove'; readonly sheet: string }
    | { readonly op: 'sheet.duplicate'; readonly sheet: string; readonly name?: string }
    | { readonly op: 'sheet.move'; readonly sheet: string; readonly index: number }
    | {
          /**
           * Sorts whole rows of `range` by `column`, moving formulas with their
           * row. `column` is a one-based offset inside the range or an absolute
           * column letter. Numbers sort before text, text before booleans, and
           * blanks always trail. The operation is rejected when the range
           * overlaps a merged range.
           */
          readonly op: 'range.sort';
          readonly sheet: string;
          readonly range: string;
          readonly column: number | string;
          /** Ascending when true, which is the default. */
          readonly ascending?: boolean;
          /** Keeps the first row of `range` in place. */
          readonly hasHeader?: boolean;
      }
    | {
          /**
           * Fills `range`, a single row or a single column, with a linear numeric
           * series. The first cell seeds the start and, when two numeric cells are
           * present, the second seeds the step. Existing styles stay in place and
           * formulas inside the range are replaced.
           */
          readonly op: 'range.fillSeries';
          readonly sheet: string;
          readonly range: string;
          /** First value; defaults to the numeric first cell of `range`. */
          readonly start?: number;
          /** Step; defaults to the difference of the first two cells, else `1`. */
          readonly step?: number;
      }
    | {
          /**
           * Moves `range` into the transposed rectangle anchored at `target`.
           * Cell contents, styles and formula sources move verbatim, while
           * relative references inside formulas are not rewritten. Merged cells
           * and destinations that overlap the source are rejected.
           */
          readonly op: 'range.transpose';
          readonly sheet: string;
          readonly range: string;
          /** Top-left cell of the destination rectangle. */
          readonly target: CellAddress;
      }
    | RangeCopyOperation
    | RangeMoveOperation
    | RangeSwapOperation
    | RangeFlipOperation
    | { readonly op: 'range.clear'; readonly sheet: string; readonly range: string };

/** Request of `Workbook.edit()`. */
export interface EditRequest {
    /** Rejects the write when the workbook moved past this revision. */
    readonly expectedRevision?: number;
    readonly operations: readonly EditOperation[];
}

/** Result of one committed edit. */
export interface EditResult {
    readonly revision: number;
    readonly workbook: Workbook;
    /** Occurrences changed by `cell.replaceText` operations. */
    readonly replacements: number;
    /** Rectangles and replacement counts recorded by the batch. */
    readonly affected?: readonly AffectedRange[];
}

/** Which field a find hit matched. */
export type FindField = 'value' | 'formula';

/** Options of `Workbook.find()`. Both channels are searched by default. */
export interface FindOptions {
    /** Match case exactly; defaults to false. */
    readonly matchCase?: boolean;
    /** Search formula sources; defaults to true. */
    readonly formulas?: boolean;
    /** Search evaluated values; defaults to true. */
    readonly values?: boolean;
}

/** Request of `Workbook.find()`. */
export interface FindRequest {
    readonly query: string;
    /** Sheet name or id; every sheet when omitted. */
    readonly sheet?: string;
    readonly options?: FindOptions;
}

/** One match, with character offsets into the searched field. */
export interface FindHit {
    readonly sheet: string;
    readonly address: CellAddress;
    readonly field: FindField;
    readonly start: number;
    readonly end: number;
    readonly text: string;
}

/** Result of `Workbook.find()`. */
export interface FindResult {
    readonly hits: readonly FindHit[];
    /** True when the engine stopped at its 500-hit cap with more matches left. */
    readonly truncated: boolean;
}
