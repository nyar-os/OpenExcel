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
    readonly renderers: Readonly<Record<RenderFormat, SupportLevel>>;
    readonly exportTargets: Readonly<Record<ExportFormat, SupportLevel>>;
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

/** One operation of `Workbook.edit()`. */
export type EditOperation =
    | {
          readonly op: 'cell.set';
          readonly sheet: string;
          readonly reference: CellAddress;
          readonly value?: CellValueInput;
          readonly formula?: string;
      }
    | { readonly op: 'sheet.add'; readonly name: string; readonly index?: number }
    | { readonly op: 'sheet.rename'; readonly sheet: string; readonly name: string }
    | { readonly op: 'sheet.remove'; readonly sheet: string }
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
}
