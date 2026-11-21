/**
 * `new OpenExcel(options?)` workbook session.
 *
 * The constructor is synchronous and opens nothing: no network, no disk scan,
 * no window, no file. Native initialisation and imports happen inside the async
 * methods, so a failure is always awaited and always carries a stable code.
 */

import { OpenExcelError } from './errors.js';
import { formatFromPath, type WorkbookFormat, type WorkbookIntent } from './formats.js';
import * as service from './service.js';
import type {
    EditRequest,
    EditResult,
    ExportOptions,
    ExportReport,
    InspectOptions,
    InspectReport,
    OpenExcelCapabilities,
    OpenExcelOptions,
    SaveOptions,
    SaveReport,
    SvgRender,
    Workbook,
    WorkbookCreateOptions,
    WorkbookInput,
    WorkbookRenderOptions,
} from './types.js';

/** Workbook session. One instance may open several workbooks. */
export class OpenExcel {
    /** BCP-47 tag used for diagnostics and default copy. */
    readonly locale: string;
    private closed = false;
    private readonly handles = new Set<number>();
    private readonly workbookHandles = new Map<string, number>();

    constructor(options: OpenExcelOptions = {}) {
        this.locale = options.locale ?? 'en';
    }

    /** Whether {@link dispose} has run. */
    get disposed(): boolean {
        return this.closed;
    }

    /**
     * Format, checker, renderer and export disclosure for the installed engine.
     * Read this instead of treating a method name as a capability.
     */
    async capabilities(): Promise<OpenExcelCapabilities> {
        this.assertActive();
        return service.capabilities();
    }

    /** Creates an empty workbook and returns its committed summary. */
    async create(options: WorkbookCreateOptions = {}): Promise<Workbook> {
        this.assertActive();
        if (options.title !== undefined && (typeof options.title !== 'string' || options.title.length === 0)) {
            throw new OpenExcelError('session/argument-invalid', '`title` must be a non-empty string');
        }
        if (options.template !== undefined && (typeof options.template !== 'string' || options.template.length === 0)) {
            throw new OpenExcelError('session/argument-invalid', '`template` must be a non-empty string');
        }
        const created = service.create({
            title: options.title,
            template: options.template,
            locale: options.locale ?? this.locale,
        });
        this.track(created.handle, created.workbook);
        return created.workbook;
    }

    /** Opens a project file, spreadsheet file or in-memory package. */
    async open(input: WorkbookInput): Promise<Workbook> {
        this.assertActive();
        assertWorkbookInput(input);
        if (!('path' in input)) {
            const key = 'bytes' in input ? 'bytes' : 'input';
            throw new OpenExcelError(
                'format/unsupported',
                `opening \`${key}\` needs the package importer; write the project to a path and pass \`path\``,
                { details: { feature: `open.${key}` } },
            );
        }
        const intent: WorkbookIntent = input.intent ?? 'edit';
        if (intent !== 'edit') {
            throw new OpenExcelError(
                'format/unsupported',
                `native projects open for editing; \`${intent}\` needs the data pipeline`,
                { details: { feature: `open.intent.${intent}`, format: this.formatOf(input.path) } },
            );
        }
        const opened = service.open(input.path);
        this.track(opened.handle, opened.workbook);
        return opened.workbook;
    }

    /** Inspects an input without keeping a workbook handle. */
    async inspect(input: WorkbookInput, options?: InspectOptions): Promise<InspectReport> {
        this.assertActive();
        const opened = service.open(assertProjectPath(input));
        try {
            return service.inspect(opened.handle, options?.checks);
        } finally {
            service.close(opened.handle);
        }
    }

    /** Writes one tracked workbook to a native project path. */
    async save(workbook: Workbook | string, options: SaveOptions): Promise<SaveReport> {
        this.assertActive();
        if (options === undefined || typeof options.target !== 'string' || options.target.length === 0) {
            throw new OpenExcelError('session/argument-invalid', '`save` needs a non-empty `target` path');
        }
        return service.save(this.handleOf(workbook), options.target, options.overwrite ?? false);
    }

    /** Reads the current summary of one tracked workbook. */
    async summary(workbook: Workbook | string): Promise<Workbook> {
        this.assertActive();
        return service.summary(this.handleOf(workbook));
    }

    /** Applies one edit request and returns the committed revision and summary. */
    async edit(workbook: Workbook | string, request: EditRequest): Promise<EditResult> {
        this.assertActive();
        assertEditRequest(request);
        return service.edit(this.handleOf(workbook), request);
    }

    /** Writes one tracked workbook to an exchange format. */
    async export(workbook: Workbook | string, options: ExportOptions): Promise<ExportReport> {
        this.assertActive();
        if (options === undefined || typeof options.target !== 'string' || options.target.length === 0) {
            throw new OpenExcelError('session/argument-invalid', '`export` needs a non-empty `target` path');
        }
        return service.exportWorkbook(this.handleOf(workbook), options.format, options.target, options.overwrite ?? false);
    }

    /** Renders one sheet of a tracked workbook as SVG. */
    async render(workbook: Workbook | string, options: WorkbookRenderOptions = {}): Promise<SvgRender> {
        this.assertActive();
        const format = options.format ?? 'svg';
        if (format !== 'svg') {
            throw new OpenExcelError('format/unsupported', `\`${format}\` rendering is not available in this build`, {
                details: { feature: `render.${format}` },
            });
        }
        return service.render(this.handleOf(workbook), options);
    }

    /** Releases the session and every workbook handle it still tracks. Idempotent. */
    async dispose(): Promise<void> {
        if (this.closed) {
            return;
        }
        this.closed = true;
        for (const handle of this.handles) {
            try {
                service.close(handle);
            } catch {
                // The addon may already be gone; dispose still resolves.
            }
        }
        this.handles.clear();
        this.workbookHandles.clear();
    }

    private track(handle: number, workbook: Workbook): void {
        this.handles.add(handle);
        this.workbookHandles.set(workbook.id, handle);
    }

    private handleOf(workbook: Workbook | string): number {
        const id = typeof workbook === 'string' ? workbook : workbook.id;
        const handle = this.workbookHandles.get(id);
        if (handle === undefined) {
            throw new OpenExcelError('workbook/not-found', `workbook \`${id}\` is not open in this session`);
        }
        return handle;
    }

    /** @internal Throws once the session is gone. */
    assertActive(): void {
        if (this.closed) {
            throw new OpenExcelError('session/disposed', 'the OpenExcel session has been disposed; create a new instance');
        }
    }

    private formatOf(path: string): WorkbookFormat | undefined {
        return formatFromPath(path);
    }
}

function assertProjectPath(input: WorkbookInput): string {
    assertWorkbookInput(input);
    if (!('path' in input)) {
        throw new OpenExcelError('format/unsupported', 'inspection needs a project path; in-memory packages are not supported yet', {
            details: { feature: 'inspect.bytes' },
        });
    }
    return input.path;
}

function assertEditRequest(request: EditRequest): void {
    if (request === undefined || !Array.isArray(request.operations) || request.operations.length === 0) {
        throw new OpenExcelError('session/argument-invalid', '`edit` needs a non-empty `operations` array');
    }
    if (
        request.expectedRevision !== undefined &&
        (!Number.isInteger(request.expectedRevision) || request.expectedRevision < 0)
    ) {
        throw new OpenExcelError('session/argument-invalid', '`expectedRevision` must be a non-negative integer');
    }
    for (const operation of request.operations) {
        if (typeof operation !== 'object' || operation === null || typeof operation.op !== 'string') {
            throw new OpenExcelError('session/argument-invalid', 'each edit operation needs a string `op`');
        }
    }
}

function assertWorkbookInput(input: WorkbookInput): void {
    const candidate = input as unknown as Record<string, unknown>;
    const present = (['path', 'bytes'] as const).filter((key) => candidate[key] !== undefined);
    if (present.length !== 1) {
        throw new OpenExcelError('session/argument-invalid', 'WorkbookInput must contain exactly one of `path` or `bytes`');
    }
    const value = candidate[present[0]];
    if (present[0] === 'bytes') {
        if (!(value instanceof Uint8Array)) {
            throw new OpenExcelError('session/argument-invalid', '`bytes` must be a Uint8Array');
        }
        return;
    }
    if (typeof value !== 'string' || value.length === 0) {
        throw new OpenExcelError('session/argument-invalid', `\`${present[0]}\` must be a non-empty string`);
    }
}
