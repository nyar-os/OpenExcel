/**
 * Orthogonal native operations.
 *
 * One function performs one operation and returns data; argument parsing, help
 * text and exit codes belong to `cli.ts` and to the session surface. Nothing
 * here knows about a command line.
 */

import {
    NotImplementedError,
    OpenExcelError,
    errorCodes,
    renderErrorCopy,
    type OpenExcelErrorCode,
    type OpenExcelErrorParams,
} from './errors.js';
import { nativeBinding } from './native.js';
import type {
    CellDetail,
    CellQuery,
    EditRequest,
    EditResult,
    ExportFormat,
    ExportReport,
    FindRequest,
    FindResult,
    InspectCheck,
    InspectReport,
    OpenExcelCapabilities,
    RangeMatrix,
    RangeQuery,
    RecalcReport,
    SaveReport,
    SvgRender,
    Workbook,
    WorkbookRenderOptions,
} from './types.js';

/** Result of creating or opening a workbook. */
export interface WorkbookHandle {
    readonly handle: number;
    readonly workbook: Workbook;
}

/** Arguments of {@link create}. */
export interface CreateWorkbookRequest {
    readonly title?: string;
    readonly template?: string;
    readonly locale?: string;
}

const knownCodes = new Set<string>(errorCodes);

function binding(feature: string) {
    const native = nativeBinding();
    if (!native) {
        throw new NotImplementedError(feature);
    }
    return native;
}

function decode<T>(work: () => string): T {
    try {
        return JSON.parse(work()) as T;
    } catch (error) {
        throw translate(error);
    }
}

function translate(error: unknown): OpenExcelError {
    if (error instanceof OpenExcelError) {
        return error;
    }
    const native = error instanceof Error ? (error as Error & { code?: unknown; detail?: unknown }) : undefined;
    const message = native ? native.message : String(error);
    const rawCode = native?.code;
    const code = typeof rawCode === 'string' ? rawCode : undefined;
    if (code !== undefined) {
        const params = readParams(native?.detail);
        return new OpenExcelError(code as OpenExcelErrorCode, renderErrorCopy(code, params, message), { params, cause: error });
    }
    // Compatibility entry for addons older than the structured contract, which
    // concatenated `code: message` into the Error text. Current bindings always
    // carry `code` and `detail.params`, so this split never runs against them.
    const separator = message.indexOf(': ');
    const legacy = separator > 0 ? message.slice(0, separator) : '';
    if (knownCodes.has(legacy)) {
        return new OpenExcelError(legacy as OpenExcelErrorCode, message.slice(separator + 2), { cause: error });
    }
    return new OpenExcelError('binding/failure', message, { cause: error });
}

function readParams(detail: unknown): OpenExcelErrorParams {
    if (typeof detail === 'object' && detail !== null) {
        const params = (detail as { params?: unknown }).params;
        if (typeof params === 'object' && params !== null && !Array.isArray(params)) {
            return params as OpenExcelErrorParams;
        }
    }
    return {};
}

/** Reads the capability matrix of the installed addon. */
export function capabilities(): OpenExcelCapabilities {
    return decode<OpenExcelCapabilities>(() => binding('capabilities').capabilities());
}

/** Creates a workbook and returns its handle. */
export function create(request: CreateWorkbookRequest): WorkbookHandle {
    return decode<WorkbookHandle>(() =>
        binding('create').workbook_create(
            JSON.stringify({ title: request.title, template: request.template, locale: request.locale }),
        ),
    );
}

/** Opens a project file and returns its handle. */
export function open(path: string): WorkbookHandle {
    return decode<WorkbookHandle>(() => binding('open').workbook_open(path));
}

/** Writes one open workbook to a project path. */
export function save(handle: number, target: string, overwrite: boolean): SaveReport {
    return decode<SaveReport>(() => binding('workbook.save').workbook_save(handle, target, overwrite));
}

/** Runs the rule checks of one open workbook. */
export function inspect(handle: number, checks?: readonly InspectCheck[]): InspectReport {
    return decode<InspectReport>(() =>
        checks === undefined
            ? binding('workbook.inspect').workbook_inspect(handle)
            : binding('workbook.inspect').workbook_inspect(handle, JSON.stringify(checks)),
    );
}

/** Returns the summary of one open workbook. */
export function summary(handle: number): Workbook {
    return decode<Workbook>(() => binding('workbook.summary').workbook_summary(handle));
}

/** Applies an edit request to one open workbook. */
export function edit(handle: number, request: EditRequest): EditResult {
    return decode<EditResult>(() => binding('workbook.edit').workbook_edit(handle, JSON.stringify(request)));
}

/** Reads one cell with its evaluated value. */
export function cell(handle: number, query: CellQuery): CellDetail {
    return decode<CellDetail>(() => binding('workbook.cell').workbook_cell(handle, JSON.stringify(query)));
}

/** Reads a rectangle of cells with their evaluated values. */
export function range(handle: number, query: RangeQuery): RangeMatrix {
    return decode<RangeMatrix>(() => binding('workbook.range').workbook_range(handle, JSON.stringify(query)));
}

/** Recomputes every formula of one open workbook. */
export function recalc(handle: number): RecalcReport {
    return decode<RecalcReport>(() => binding('workbook.recalc').workbook_recalc(handle));
}

/** Searches evaluated values and formula sources of one open workbook. */
export function find(handle: number, request: FindRequest): FindResult {
    return decode<FindResult>(() => binding('workbook.find').workbook_find(handle, JSON.stringify(request)));
}

/** Renders one sheet of an open workbook as SVG. */
export function render(handle: number, options: WorkbookRenderOptions = {}): SvgRender {
    return decode<SvgRender>(() => binding('workbook.render').workbook_render(handle, JSON.stringify(options)));
}

/** Writes one open workbook to an export target. */
export function exportWorkbook(
    handle: number,
    format: ExportFormat,
    target: string,
    overwrite: boolean,
): ExportReport {
    return decode<ExportReport>(() => binding('workbook.export').workbook_export(handle, format, target, overwrite));
}

/** Releases one open workbook handle. */
export function close(handle: number): void {
    binding('workbook.close').workbook_close(handle);
}
