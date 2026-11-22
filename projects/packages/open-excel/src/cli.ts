#!/usr/bin/env node
/**
 * Command-line assembly only.
 *
 * Every command maps options onto one method of `new OpenExcel()`; workbook
 * logic, saving and export stay native. Help text and errors come from
 * `@wae/commander`, so the CLI does not grow its own parser.
 */

import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { createCli } from '@wae/commander';

import { OpenExcel } from './open-excel.js';
import { runGui } from './gui.js';
import type {
    CellDetail,
    CellStyleInput,
    EditOperation,
    ExportFormat,
    FindResult,
    FreezeCommandOptions,
    InspectCheck,
    InspectReport,
    OpenExcelCapabilities,
    RangeMatrix,
    RecalcReport,
    SheetVisibilityCommandOptions,
    Workbook,
} from './types.js';

const here = path.dirname(fileURLToPath(import.meta.url));

const cli = createCli('open-excel').locales(path.join(here, '..', 'locales')).intro('cli.intro');

/** Runs `work` against one session and always releases its handles. */
async function withSession<T>(work: (excel: OpenExcel) => Promise<T>): Promise<T> {
    const excel = new OpenExcel();
    try {
        return await work(excel);
    } finally {
        await excel.dispose();
    }
}

// Exported for contract audits that import the built entry to inspect the
// command tree. Nothing runs until the file is executed as the program entry.
export { cli };

interface ParsedOptions {
    readonly _?: readonly string[];
    readonly [key: string]: unknown;
}

const text = (options: ParsedOptions, key: string): string | undefined =>
    typeof options[key] === 'string' ? (options[key] as string) : undefined;

const flag = (options: ParsedOptions, key: string): boolean => options[key] === true || options[key] === 'true';

const input = (options: ParsedOptions): string | undefined => options._?.[0];

/** Export targets the CLI accepts; each still reports its real support level. */
const exportTargets: readonly ExportFormat[] = ['xlsx', 'xlsm', 'csv', 'tsv', 'json', 'html', 'pdf'];

function describeCapabilities(matrix: OpenExcelCapabilities): string {
    const formats = Object.entries(matrix.formats)
        .filter(([, entry]) => entry.import === 'supported')
        .map(([name]) => name);
    const checks = Object.entries(matrix.checks)
        .filter(([, level]) => level === 'supported')
        .map(([name]) => name);
    const targets = Object.entries(matrix.exportTargets)
        .filter(([, level]) => level === 'supported' || level === 'partial')
        .map(([name, level]) => (level === 'partial' ? `${name} (partial)` : name));
    return [
        `engine: ${matrix.engine.package} ${matrix.engine.version}`,
        `project: ${matrix.project}`,
        `workbook formats: ${formats.length > 0 ? formats.join(', ') : 'none supported'}`,
        `checks: ${checks.length > 0 ? checks.join(', ') : 'none supported'}`,
        `export targets: ${targets.length > 0 ? targets.join(', ') : 'none supported'}`,
    ].join('\n');
}

function describeWorkbook(workbook: Workbook): string {
    const lines = [`workbook ${workbook.id} revision=${workbook.revision} sheets=${workbook.sheets.length} cells=${workbook.cells}`];
    for (const sheet of workbook.sheets) {
        lines.push(`  [${sheet.index}] ${sheet.id} ${sheet.name}${sheet.hidden ? ' (hidden)' : ''} cells=${sheet.cells}`);
    }
    return lines.join('\n');
}

function describeReport(report: InspectReport): string {
    const lines = [`inspect revision=${report.revision} checker=${report.checkerVersion} passed=${report.passed}`];
    for (const issue of report.issues) {
        lines.push(`  ${issue.severity} ${issue.code} ${issue.sheetId ?? '-'} ${issue.cell ?? '-'} ${issue.message}`);
    }
    return lines.join('\n');
}

function describeCell(detail: CellDetail): string {
    const parts = [`${detail.sheet}!${detail.address} kind=${detail.kind}`];
    if (detail.formula !== null) {
        parts.push(`formula=${detail.formula}`);
    }
    parts.push(`value=${detail.value === null ? '(empty)' : String(detail.value)}`);
    if (detail.cached !== null) {
        parts.push(`cached=${String(detail.cached)}`);
    }
    return parts.join(' ');
}

function describeRange(matrix: RangeMatrix): string {
    const lines = [`range ${matrix.sheet}!${matrix.range} rows=${matrix.rows.length}`];
    for (const row of matrix.rows) {
        lines.push(row.map((cell) => (cell.value === null ? '' : String(cell.value))).join('\t'));
    }
    return lines.join('\n');
}

function describeRecalc(report: RecalcReport): string {
    return `recalc revision=${report.revision} changed=${report.changed}`;
}

function describeFind(result: FindResult): string {
    const lines = [`find hits=${result.hits.length} truncated=${result.truncated}`];
    for (const hit of result.hits) {
        lines.push(`  ${hit.sheet}!${hit.address} ${hit.field} ${hit.start}-${hit.end} ${hit.text}`);
    }
    return lines.join('\n');
}

function parseChecks(value: string | undefined): InspectCheck[] | undefined {
    if (value === undefined) {
        return undefined;
    }
    return value
        .split(',')
        .map((name) => name.trim())
        .filter((name) => name.length > 0) as InspectCheck[];
}

function parseValue(value: string): string | number | boolean | null {
    const trimmed = value.trim();
    if (trimmed === 'null') {
        return null;
    }
    if (trimmed === 'true') {
        return true;
    }
    if (trimmed === 'false') {
        return false;
    }
    if (trimmed.length > 0 && /^-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
        return Number(trimmed);
    }
    return value;
}

/** Parses `--ops`; `undefined` means the payload is not a non-empty JSON array. */
function parseOperations(value: string): EditOperation[] | undefined {
    try {
        const parsed: unknown = JSON.parse(value);
        if (!Array.isArray(parsed) || parsed.length === 0) {
            return undefined;
        }
        return parsed as EditOperation[];
    } catch {
        return undefined;
    }
}

cli.command('capabilities', 'cli.cmd.capabilities').actionCtx(async (_options, _args, ctx) => {
    const matrix = await withSession((excel) => excel.capabilities());
    ctx.emitResult(matrix, describeCapabilities(matrix));
    return 0;
});

cli
    .command('create', 'cli.cmd.create')
    .option('--out <file>', 'cli.opt.out')
    .option('--title <title>', 'cli.opt.title')
    .option('--template <name>', 'cli.opt.template')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const out = text(options, 'out');
        if (out === undefined) {
            ctx.fail('open-excel.cli.out-required', { message: '`--out <file>` is required' });
            return 2;
        }
        await withSession(async (excel) => {
            const workbook = await excel.create({ title: text(options, 'title'), template: text(options, 'template') });
            const report = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
            ctx.emitResult(
                { ...report, workbook },
                `created ${out} revision=${report.revision} sheets=${report.sheets}`,
            );
        });
        return 0;
    });

cli.command('summary', 'cli.cmd.summary').actionCtx(async (options, _args, ctx) => {
    const file = input(options);
    if (file === undefined) {
        ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
        return 2;
    }
    await withSession(async (excel) => {
        const workbook = await excel.open({ path: file });
        const summary = await excel.summary(workbook);
        ctx.emitResult(summary, describeWorkbook(summary));
    });
    return 0;
});

cli
    .command('save', 'cli.cmd.save')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const out = text(options, 'out') ?? file;
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const report = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
            ctx.emitResult(report, `saved ${out} revision=${report.revision} sheets=${report.sheets}`);
        });
        return 0;
    });

cli
    .command('edit', 'cli.cmd.edit')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--cell <ref>', 'cli.opt.cell')
    .option('--value <text>', 'cli.opt.value')
    .option('--formula <formula>', 'cli.opt.formula')
    .option('--ops <json>', 'cli.opt.ops')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const rawOperations = text(options, 'ops');
        const reference = text(options, 'cell');
        const formula = text(options, 'formula');
        const rawValue = text(options, 'value');
        let operations: EditOperation[];
        let label: string;
        if (rawOperations !== undefined) {
            if (reference !== undefined || formula !== undefined || rawValue !== undefined) {
                ctx.fail('open-excel.cli.ops-exclusive', {
                    message: '`--ops` cannot be combined with `--cell`, `--value` or `--formula`',
                });
                return 2;
            }
            const parsed = parseOperations(rawOperations);
            if (parsed === undefined) {
                ctx.fail('open-excel.cli.ops-invalid', {
                    message: '`--ops` must be a JSON array of operations',
                });
                return 2;
            }
            operations = parsed;
            label = `${operations.length} operation${operations.length === 1 ? '' : 's'}`;
        }
        else {
            if (reference === undefined) {
                ctx.fail('open-excel.cli.cell-required', { message: '`--cell <A1>` is required' });
                return 2;
            }
            if (formula === undefined && rawValue === undefined) {
                ctx.fail('open-excel.cli.value-required', { message: '`--value` or `--formula` is required' });
                return 2;
            }
            operations = [
                {
                    op: 'cell.set',
                    sheet: text(options, 'sheet') ?? 'Sheet1',
                    reference,
                    ...(formula === undefined ? { value: parseValue(rawValue ?? '') } : { formula }),
                },
            ];
            label = reference;
        }
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult({ ...result, saved }, `edited ${label} revision=${saved.revision} saved=${out}`);
                return;
            }
            ctx.emitResult(result, `edited ${label} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('sort', 'cli.cmd.sort')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--range <range>', 'cli.opt.range')
    .option('--column <column>', 'cli.opt.column')
    .option('--descending', 'cli.opt.descending')
    .option('--header', 'cli.opt.header')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const range = text(options, 'range');
        if (range === undefined) {
            ctx.fail('open-excel.cli.range-required', { message: '`--range <A1:C4>` is required' });
            return 2;
        }
        const rawColumn = text(options, 'column');
        if (rawColumn === undefined || rawColumn.trim().length === 0) {
            ctx.fail('open-excel.cli.column-required', { message: '`--column <n|letter>` is required' });
            return 2;
        }
        const trimmedColumn = rawColumn.trim();
        const column = /^\d+$/.test(trimmedColumn) ? Number(trimmedColumn) : trimmedColumn;
        const sheet = text(options, 'sheet') ?? 'Sheet1';
        const operation: EditOperation = {
            op: 'range.sort',
            sheet,
            range,
            column,
            ascending: !flag(options, 'descending'),
            hasHeader: flag(options, 'header'),
        };
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult(
                    { ...result, saved },
                    `sorted ${sheet}!${range} revision=${saved.revision} saved=${out}`,
                );
                return;
            }
            ctx.emitResult(result, `sorted ${sheet}!${range} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('copy', 'cli.cmd.copy')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--range <range>', 'cli.opt.range')
    .option('--to <cell>', 'cli.opt.to-target')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const range = text(options, 'range');
        if (range === undefined) {
            ctx.fail('open-excel.cli.range-required', { message: '`--range <A1:C4>` is required' });
            return 2;
        }
        const target = text(options, 'to');
        if (target === undefined || target.trim().length === 0) {
            ctx.fail('open-excel.cli.target-required', { message: '`--to <cell>` is required' });
            return 2;
        }
        const sheet = text(options, 'sheet') ?? 'Sheet1';
        const operation: EditOperation = { op: 'range.copy', sheet, range, target };
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult(
                    { ...result, saved },
                    `copied ${sheet}!${range} to ${target} revision=${saved.revision} saved=${out}`,
                );
                return;
            }
            ctx.emitResult(result, `copied ${sheet}!${range} to ${target} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('move', 'cli.cmd.move')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--range <range>', 'cli.opt.range')
    .option('--to <cell>', 'cli.opt.to-move-target')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const range = text(options, 'range');
        if (range === undefined) {
            ctx.fail('open-excel.cli.range-required', { message: '`--range <A1:C4>` is required' });
            return 2;
        }
        const target = text(options, 'to');
        if (target === undefined || target.trim().length === 0) {
            ctx.fail('open-excel.cli.target-required', { message: '`--to <cell>` is required' });
            return 2;
        }
        const sheet = text(options, 'sheet') ?? 'Sheet1';
        const operation: EditOperation = { op: 'range.move', sheet, range, target };
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult(
                    { ...result, saved },
                    `moved ${sheet}!${range} to ${target} revision=${saved.revision} saved=${out}`,
                );
                return;
            }
            ctx.emitResult(result, `moved ${sheet}!${range} to ${target} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('format', 'cli.cmd.format')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--range <range>', 'cli.opt.range')
    .option('--bold', 'cli.opt.bold')
    .option('--italic', 'cli.opt.italic')
    .option('--font-size <pt>', 'cli.opt.font-size')
    .option('--font-color <color>', 'cli.opt.font-color')
    .option('--fill <color>', 'cli.opt.fill')
    .option('--number-format <format>', 'cli.opt.number-format')
    .option('--align <align>', 'cli.opt.align')
    .option('--wrap', 'cli.opt.wrap')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const range = text(options, 'range');
        if (range === undefined) {
            ctx.fail('open-excel.cli.range-required', { message: '`--range <A1:C4>` is required' });
            return 2;
        }
        const style: {
            bold?: boolean;
            italic?: boolean;
            fontSize?: number;
            fontColor?: string;
            fill?: string;
            numberFormat?: string;
            align?: CellStyleInput['align'];
            wrap?: boolean;
        } = {};
        if (flag(options, 'bold')) {
            style.bold = true;
        }
        if (flag(options, 'italic')) {
            style.italic = true;
        }
        if (flag(options, 'wrap')) {
            style.wrap = true;
        }
        const fontSize = text(options, 'font-size');
        if (fontSize !== undefined) {
            const value = Number(fontSize);
            if (!Number.isFinite(value) || value <= 0) {
                ctx.fail('open-excel.cli.font-size-invalid', {
                    message: '`--font-size` must be a positive number',
                });
                return 2;
            }
            style.fontSize = value;
        }
        const fontColor = text(options, 'font-color');
        if (fontColor !== undefined) {
            style.fontColor = fontColor;
        }
        const fill = text(options, 'fill');
        if (fill !== undefined) {
            style.fill = fill;
        }
        const numberFormat = text(options, 'number-format');
        if (numberFormat !== undefined) {
            style.numberFormat = numberFormat;
        }
        const align = text(options, 'align');
        if (align !== undefined) {
            const allowed = ['left', 'center', 'right', 'justify', 'distribute', 'fill'] as const;
            if (!(allowed as readonly string[]).includes(align)) {
                ctx.fail('open-excel.cli.align-invalid', {
                    message: `\`--align\` must be one of ${allowed.join(', ')}`,
                });
                return 2;
            }
            style.align = align as CellStyleInput['align'];
        }
        if (Object.keys(style).length === 0) {
            ctx.fail('open-excel.cli.style-required', { message: 'at least one style option is required' });
            return 2;
        }
        const sheet = text(options, 'sheet') ?? 'Sheet1';
        const operation: EditOperation = { op: 'format.set', sheet, range, style };
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult(
                    { ...result, saved },
                    `formatted ${sheet}!${range} revision=${saved.revision} saved=${out}`,
                );
                return;
            }
            ctx.emitResult(result, `formatted ${sheet}!${range} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('layout', 'cli.cmd.layout')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--column <column>', 'cli.opt.layout-column')
    .option('--width <units>', 'cli.opt.width')
    .option('--row <row>', 'cli.opt.layout-row')
    .option('--height <pt>', 'cli.opt.height')
    .option('--hide', 'cli.opt.hide')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const rawWidth = text(options, 'width');
        const rawHeight = text(options, 'height');
        if (rawWidth === undefined && rawHeight === undefined) {
            ctx.fail('open-excel.cli.layout-size-required', {
                message: '`--width <units>` or `--height <pt>` is required',
            });
            return 2;
        }
        const sheet = text(options, 'sheet') ?? 'Sheet1';
        const hide = flag(options, 'hide');
        const operations: EditOperation[] = [];
        const labels: string[] = [];
        if (rawWidth !== undefined) {
            const rawColumn = text(options, 'column')?.trim();
            if (rawColumn === undefined || rawColumn.length === 0) {
                ctx.fail('open-excel.cli.column-required', {
                    message: '`--width` needs `--column <letter|n>`',
                });
                return 2;
            }
            const width = Number(rawWidth);
            if (!Number.isFinite(width)) {
                ctx.fail('open-excel.cli.layout-width-invalid', { message: '`--width` must be a number' });
                return 2;
            }
            const column = /^\d+$/.test(rawColumn) ? Number(rawColumn) : rawColumn;
            operations.push({ op: 'column.setWidth', sheet, column, width, ...(hide ? { hidden: true } : {}) });
            labels.push(`column ${rawColumn}`);
        }
        if (rawHeight !== undefined) {
            const rawRow = text(options, 'row')?.trim();
            if (rawRow === undefined || !/^\d+$/.test(rawRow) || Number(rawRow) < 1) {
                ctx.fail('open-excel.cli.row-required', { message: '`--height` needs a positive `--row <n>`' });
                return 2;
            }
            const height = Number(rawHeight);
            if (!Number.isFinite(height)) {
                ctx.fail('open-excel.cli.layout-height-invalid', { message: '`--height` must be a number' });
                return 2;
            }
            operations.push({ op: 'row.setHeight', sheet, row: Number(rawRow), height, ...(hide ? { hidden: true } : {}) });
            labels.push(`row ${rawRow}`);
        }
        const label = labels.join(' and ');
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult({ ...result, saved }, `laid out ${label} revision=${saved.revision} saved=${out}`);
                return;
            }
            ctx.emitResult(result, `laid out ${label} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('row', 'cli.cmd.row')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--insert <row>', 'cli.opt.insert-row')
    .option('--delete <row>', 'cli.opt.delete-row')
    .option('--count <n>', 'cli.opt.count')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const insert = text(options, 'insert');
        const remove = text(options, 'delete');
        if ((insert === undefined) === (remove === undefined)) {
            ctx.fail('open-excel.cli.row-action-required', {
                message: 'exactly one of `--insert <row>` or `--delete <row>` is required',
            });
            return 2;
        }
        const rawRow = (insert ?? remove ?? '').trim();
        const row = /^\d+$/.test(rawRow) ? Number(rawRow) : Number.NaN;
        if (!Number.isInteger(row) || row < 1 || row > 1_048_576) {
            ctx.fail('open-excel.cli.row-invalid', {
                message: '`--insert`/`--delete` must be a one-based row between 1 and 1048576',
            });
            return 2;
        }
        const rawCount = text(options, 'count');
        const count = rawCount === undefined ? 1 : /^\d+$/.test(rawCount.trim()) ? Number(rawCount.trim()) : Number.NaN;
        if (!Number.isInteger(count) || count < 1) {
            ctx.fail('open-excel.cli.count-invalid', { message: '`--count` must be a positive integer' });
            return 2;
        }
        if (row + count - 1 > 1_048_576) {
            ctx.fail('open-excel.cli.row-band-invalid', {
                message: `rows ${row}..${row + count - 1} leave the sheet at row 1048576`,
            });
            return 2;
        }
        const sheet = text(options, 'sheet') ?? 'Sheet1';
        const operation: EditOperation =
            insert !== undefined
                ? { op: 'row.insert', sheet, row, count }
                : { op: 'row.delete', sheet, row, count };
        const label = `${insert !== undefined ? 'inserted' : 'deleted'} ${count} row${count === 1 ? '' : 's'} at ${row} in ${sheet}`;
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult({ ...result, saved }, `${label} revision=${saved.revision} saved=${out}`);
                return;
            }
            ctx.emitResult(result, `${label} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('column', 'cli.cmd.column')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--insert <column>', 'cli.opt.insert-column')
    .option('--delete <column>', 'cli.opt.delete-column')
    .option('--count <n>', 'cli.opt.count')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const insert = text(options, 'insert');
        const remove = text(options, 'delete');
        if ((insert === undefined) === (remove === undefined)) {
            ctx.fail('open-excel.cli.column-action-required', {
                message: 'exactly one of `--insert <column>` or `--delete <column>` is required',
            });
            return 2;
        }
        const rawColumn = (insert ?? remove ?? '').trim();
        let column: number | string;
        if (/^\d+$/.test(rawColumn)) {
            column = Number(rawColumn);
            if (column < 1 || column > 16_384) {
                ctx.fail('open-excel.cli.column-invalid', {
                    message: '`--insert`/`--delete` must be a one-based column between 1 and 16384 or a letter',
                });
                return 2;
            }
        }
        else if (/^[A-Za-z]{1,3}$/.test(rawColumn)) {
            column = rawColumn;
        }
        else {
            ctx.fail('open-excel.cli.column-invalid', {
                message: '`--insert`/`--delete` must be a one-based column between 1 and 16384 or a letter',
            });
            return 2;
        }
        const rawCount = text(options, 'count');
        const count = rawCount === undefined ? 1 : /^\d+$/.test(rawCount.trim()) ? Number(rawCount.trim()) : Number.NaN;
        if (!Number.isInteger(count) || count < 1) {
            ctx.fail('open-excel.cli.count-invalid', { message: '`--count` must be a positive integer' });
            return 2;
        }
        if (typeof column === 'number' && column + count - 1 > 16_384) {
            ctx.fail('open-excel.cli.column-band-invalid', {
                message: `columns ${column}..${column + count - 1} leave the sheet at column 16384`,
            });
            return 2;
        }
        const sheet = text(options, 'sheet') ?? 'Sheet1';
        const operation: EditOperation =
            insert !== undefined
                ? { op: 'column.insert', sheet, column, count }
                : { op: 'column.delete', sheet, column, count };
        const label = `${insert !== undefined ? 'inserted' : 'deleted'} ${count} column${count === 1 ? '' : 's'} at ${rawColumn} in ${sheet}`;
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult({ ...result, saved }, `${label} revision=${saved.revision} saved=${out}`);
                return;
            }
            ctx.emitResult(result, `${label} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('freeze', 'cli.cmd.freeze')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--at <cell>', 'cli.opt.at')
    .option('--rows <n>', 'cli.opt.rows')
    .option('--columns <n>', 'cli.opt.columns')
    .option('--clear', 'cli.opt.clear')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const rawRows = text(options, 'rows');
        const rawColumns = text(options, 'columns');
        const requested: FreezeCommandOptions = {
            sheet: text(options, 'sheet'),
            at: text(options, 'at'),
            rows: rawRows === undefined ? undefined : Number(rawRows),
            columns: rawColumns === undefined ? undefined : Number(rawColumns),
            clear: flag(options, 'clear') || undefined,
        };
        const numeric = requested.rows !== undefined || requested.columns !== undefined;
        if ([requested.at !== undefined, numeric, requested.clear === true].filter(Boolean).length !== 1) {
            ctx.fail('open-excel.cli.freeze-action-required', {
                message: 'exactly one of `--at <cell>`, `--rows <n> --columns <m>` or `--clear` is required',
            });
            return 2;
        }
        const sheet = requested.sheet ?? 'Sheet1';
        let operation: EditOperation;
        let label: string;
        if (requested.clear === true) {
            operation = { op: 'sheet.freeze', sheet, clear: true };
            label = `cleared the freeze on ${sheet}`;
        }
        else if (requested.at !== undefined) {
            operation = { op: 'sheet.freeze', sheet, cell: requested.at };
            label = `froze ${sheet} at ${requested.at}`;
        }
        else {
            const { rows, columns } = requested;
            if (!Number.isInteger(rows) || (rows ?? 0) < 0 || !Number.isInteger(columns) || (columns ?? 0) < 0) {
                ctx.fail('open-excel.cli.freeze-numbers-invalid', {
                    message: '`--rows` and `--columns` must both be non-negative integers',
                });
                return 2;
            }
            operation = { op: 'sheet.freeze', sheet, rows: rows as number, columns: columns as number };
            label = `froze ${rows} rows and ${columns} columns on ${sheet}`;
        }
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult({ ...result, saved }, `${label} revision=${saved.revision} saved=${out}`);
                return;
            }
            ctx.emitResult(result, `${label} revision=${result.revision}`);
        });
        return 0;
    });

cli
    .command('sheet', 'cli.cmd.sheet')
    .option('--add <name>', 'cli.opt.add')
    .option('--rename <name>', 'cli.opt.rename')
    .option('--to <name>', 'cli.opt.to')
    .option('--remove <name>', 'cli.opt.remove')
    .option('--duplicate <name>', 'cli.opt.duplicate')
    .option('--as <name>', 'cli.opt.as')
    .option('--move <name>', 'cli.opt.move')
    .option('--index <n>', 'cli.opt.index')
    .option('--hide <name>', 'cli.opt.hide-sheet')
    .option('--show <name>', 'cli.opt.show-sheet')
    .option('--very-hidden <name>', 'cli.opt.very-hidden-sheet')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const add = text(options, 'add');
        const rename = text(options, 'rename');
        const remove = text(options, 'remove');
        const duplicate = text(options, 'duplicate');
        const move = text(options, 'move');
        const visibility: SheetVisibilityCommandOptions = {
            hide: text(options, 'hide'),
            show: text(options, 'show'),
            veryHidden: text(options, 'very-hidden'),
        };
        const selected = [add, rename, remove, duplicate, move, visibility.hide, visibility.show, visibility.veryHidden]
            .filter((value) => value !== undefined);
        if (selected.length !== 1) {
            ctx.fail('open-excel.cli.sheet-action-required', {
                message:
                    'exactly one of `--add`, `--rename`, `--remove`, `--duplicate`, `--move`, `--hide`, `--show` or `--very-hidden` is required',
            });
            return 2;
        }
        let operation: EditOperation;
        let label: string;
        if (add !== undefined) {
            operation = { op: 'sheet.add', name: add };
            label = `added ${add}`;
        }
        else if (rename !== undefined) {
            const to = text(options, 'to');
            if (to === undefined) {
                ctx.fail('open-excel.cli.to-required', { message: '`--rename` needs `--to <new>`' });
                return 2;
            }
            operation = { op: 'sheet.rename', sheet: rename, name: to };
            label = `renamed ${rename} to ${to}`;
        }
        else if (remove !== undefined) {
            operation = { op: 'sheet.remove', sheet: remove };
            label = `removed ${remove}`;
        }
        else if (duplicate !== undefined) {
            const name = text(options, 'as');
            if (name === undefined) {
                ctx.fail('open-excel.cli.as-required', { message: '`--duplicate` needs `--as <new>`' });
                return 2;
            }
            operation = { op: 'sheet.duplicate', sheet: duplicate, name };
            label = `duplicated ${duplicate} as ${name}`;
        }
        else if (visibility.hide !== undefined) {
            operation = { op: 'sheet.setVisibility', sheet: visibility.hide, visibility: 'hidden' };
            label = `hid ${visibility.hide}`;
        }
        else if (visibility.show !== undefined) {
            operation = { op: 'sheet.setVisibility', sheet: visibility.show, visibility: 'visible' };
            label = `showed ${visibility.show}`;
        }
        else if (visibility.veryHidden !== undefined) {
            operation = { op: 'sheet.setVisibility', sheet: visibility.veryHidden, visibility: 'veryHidden' };
            label = `very-hid ${visibility.veryHidden}`;
        }
        else {
            const rawIndex = text(options, 'index');
            const index = rawIndex === undefined ? Number.NaN : Number(rawIndex);
            if (!Number.isInteger(index) || index < 0) {
                ctx.fail('open-excel.cli.index-required', { message: '`--move` needs a zero-based `--index <n>`' });
                return 2;
            }
            operation = { op: 'sheet.move', sheet: move as string, index };
            label = `moved ${move} to ${index}`;
        }
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.edit(workbook, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = await excel.save(workbook, { target: out, overwrite: flag(options, 'overwrite') });
                ctx.emitResult({ ...result, saved }, `sheet ${label} revision=${saved.revision} saved=${out}`);
                return;
            }
            ctx.emitResult(result, `sheet ${label} revision=${result.revision}`);
        });
        return 0;
    });

cli.command('open', 'cli.cmd.open').actionCtx(async (options, _args, ctx) => {
    const file = input(options);
    if (file === undefined) {
        ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
        return 2;
    }
    await withSession(async (excel) => {
        const workbook = await excel.open({ path: file });
        ctx.emitResult(workbook, describeWorkbook(workbook));
    });
    return 0;
});

cli
    .command('inspect', 'cli.cmd.inspect')
    .option('--checks <list>', 'cli.opt.checks')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        let passed = false;
        await withSession(async (excel) => {
            const report = await excel.inspect({ path: file }, { checks: parseChecks(text(options, 'checks')) });
            ctx.emitResult(report, describeReport(report));
            passed = report.passed;
        });
        return passed ? 0 : 1;
    });

cli
    .command('cell', 'cli.cmd.cell')
    .option('--sheet <name>', 'cli.opt.sheet')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const address = options._?.[1];
        if (address === undefined) {
            ctx.fail('open-excel.cli.address-required', { message: 'an A1 cell address such as `B4` is required' });
            return 2;
        }
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const detail = await excel.cell(workbook, { sheet: text(options, 'sheet'), address });
            ctx.emitResult(detail, describeCell(detail));
        });
        return 0;
    });

cli
    .command('range', 'cli.cmd.range')
    .option('--sheet <name>', 'cli.opt.sheet')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const range = options._?.[1];
        if (range === undefined) {
            ctx.fail('open-excel.cli.range-required', { message: 'an A1 range such as `A1:C4` is required' });
            return 2;
        }
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const matrix = await excel.range(workbook, { sheet: text(options, 'sheet'), range });
            ctx.emitResult(matrix, describeRange(matrix));
        });
        return 0;
    });

cli.command('recalc', 'cli.cmd.recalc').actionCtx(async (options, _args, ctx) => {
    const file = input(options);
    if (file === undefined) {
        ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
        return 2;
    }
    await withSession(async (excel) => {
        const workbook = await excel.open({ path: file });
        const report = await excel.recalc(workbook);
        ctx.emitResult(report, describeRecalc(report));
    });
    return 0;
});

cli
    .command('find', 'cli.cmd.find')
    .option('--sheet <name>', 'cli.opt.sheet')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const query = options._?.[1];
        if (query === undefined || query.trim().length === 0) {
            ctx.fail('open-excel.cli.query-required', { message: 'a non-empty search query is required' });
            return 2;
        }
        const sheet = text(options, 'sheet');
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const result = await excel.find(workbook, sheet === undefined ? { query } : { query, sheet });
            ctx.emitResult(result, describeFind(result));
        });
        return 0;
    });

cli
    .command('export', 'cli.cmd.export')
    .option('--format <name>', 'cli.opt.format')
    .option('--out <file>', 'cli.opt.out')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const format = text(options, 'format');
        if (format === undefined) {
            ctx.fail('open-excel.cli.format-required', { message: '`--format <name>` is required' });
            return 2;
        }
        if (!(exportTargets as readonly string[]).includes(format)) {
            ctx.fail('open-excel.cli.format-invalid', {
                message: `\`--format\` must be one of ${exportTargets.join(', ')}`,
            });
            return 2;
        }
        const out = text(options, 'out');
        if (out === undefined) {
            ctx.fail('open-excel.cli.out-required', { message: '`--out <file>` is required' });
            return 2;
        }
        await withSession(async (excel) => {
            const workbook = await excel.open({ path: file });
            const report = await excel.export(workbook, { format: format as ExportFormat, target: out });
            ctx.emitResult(report, `exported ${out} format=${report.format} losses=${report.losses.length}`);
        });
        return 0;
    });

cli.command('gui', 'cli.cmd.gui').actionCtx(async (_options, _args, ctx) => {
    try {
        const status = await runGui({});
        ctx.emitResult(status, `window ${status.state}`);
        return 0;
    } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        ctx.fail('open-excel.cli.gui-failed', { message });
        return 1;
    }
});

const entry = process.argv[1];
if (entry !== undefined && import.meta.url === pathToFileURL(entry).href) {
    const code = await cli.parse(process.argv);
    process.exitCode = code;
}
