#!/usr/bin/env node
/**
 * Command-line assembly only.
 *
 * Every command maps options onto one operation of `service.ts`; workbook logic,
 * saving and export stay native. Help text and errors come from
 * `@wae/commander`, so the CLI does not grow its own parser.
 */

import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { createCli } from '@wae/commander';

import * as service from './service.js';
import { runGui } from './gui.js';
import type {
    CellDetail,
    EditOperation,
    ExportFormat,
    InspectCheck,
    InspectReport,
    OpenExcelCapabilities,
    RangeMatrix,
    RecalcReport,
    Workbook,
} from './types.js';

const here = path.dirname(fileURLToPath(import.meta.url));

const cli = createCli('open-excel').locales(path.join(here, '..', 'locales')).intro('cli.intro');

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
        .filter(([, level]) => level === 'supported')
        .map(([name]) => name);
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

cli.command('capabilities', 'cli.cmd.capabilities').actionCtx(async (_options, _args, ctx) => {
    const matrix = service.capabilities();
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
        const created = service.create({ title: text(options, 'title'), template: text(options, 'template') });
        try {
            const report = service.save(created.handle, out, flag(options, 'overwrite'));
            ctx.emitResult(
                { ...report, workbook: created.workbook },
                `created ${out} revision=${report.revision} sheets=${report.sheets}`,
            );
            return 0;
        } finally {
            service.close(created.handle);
        }
    });

cli.command('summary', 'cli.cmd.summary').actionCtx(async (options, _args, ctx) => {
    const file = input(options);
    if (file === undefined) {
        ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
        return 2;
    }
    const opened = service.open(file);
    try {
        const workbook = service.summary(opened.handle);
        ctx.emitResult(workbook, describeWorkbook(workbook));
        return 0;
    } finally {
        service.close(opened.handle);
    }
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
        const opened = service.open(file);
        try {
            const report = service.save(opened.handle, out, flag(options, 'overwrite'));
            ctx.emitResult(report, `saved ${out} revision=${report.revision} sheets=${report.sheets}`);
            return 0;
        } finally {
            service.close(opened.handle);
        }
    });

cli
    .command('edit', 'cli.cmd.edit')
    .option('--sheet <name>', 'cli.opt.sheet')
    .option('--cell <ref>', 'cli.opt.cell')
    .option('--value <text>', 'cli.opt.value')
    .option('--formula <formula>', 'cli.opt.formula')
    .option('--out <file>', 'cli.opt.out')
    .option('--overwrite', 'cli.opt.overwrite')
    .actionCtx(async (options, _args, ctx) => {
        const file = input(options);
        if (file === undefined) {
            ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
            return 2;
        }
        const reference = text(options, 'cell');
        if (reference === undefined) {
            ctx.fail('open-excel.cli.cell-required', { message: '`--cell <A1>` is required' });
            return 2;
        }
        const formula = text(options, 'formula');
        const rawValue = text(options, 'value');
        if (formula === undefined && rawValue === undefined) {
            ctx.fail('open-excel.cli.value-required', { message: '`--value` or `--formula` is required' });
            return 2;
        }
        const operation: EditOperation = {
            op: 'cell.set',
            sheet: text(options, 'sheet') ?? 'Sheet1',
            reference,
            ...(formula === undefined ? { value: parseValue(rawValue ?? '') } : { formula }),
        };
        const opened = service.open(file);
        try {
            const result = service.edit(opened.handle, { operations: [operation] });
            const out = text(options, 'out');
            if (out !== undefined) {
                const saved = service.save(opened.handle, out, flag(options, 'overwrite'));
                ctx.emitResult({ ...result, saved }, `edited ${reference} revision=${saved.revision} saved=${out}`);
                return 0;
            }
            ctx.emitResult(result, `edited ${reference} revision=${result.revision}`);
            return 0;
        } finally {
            service.close(opened.handle);
        }
    });

cli.command('open', 'cli.cmd.open').actionCtx(async (options, _args, ctx) => {
    const file = input(options);
    if (file === undefined) {
        ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
        return 2;
    }
    const opened = service.open(file);
    try {
        ctx.emitResult(opened.workbook, describeWorkbook(opened.workbook));
        return 0;
    } finally {
        service.close(opened.handle);
    }
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
        const opened = service.open(file);
        try {
            const report = service.inspect(opened.handle, parseChecks(text(options, 'checks')));
            ctx.emitResult(report, describeReport(report));
            return report.passed ? 0 : 1;
        } finally {
            service.close(opened.handle);
        }
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
        const opened = service.open(file);
        try {
            const detail = service.cell(opened.handle, { sheet: text(options, 'sheet'), address });
            ctx.emitResult(detail, describeCell(detail));
            return 0;
        } finally {
            service.close(opened.handle);
        }
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
        const opened = service.open(file);
        try {
            const matrix = service.range(opened.handle, { sheet: text(options, 'sheet'), range });
            ctx.emitResult(matrix, describeRange(matrix));
            return 0;
        } finally {
            service.close(opened.handle);
        }
    });

cli.command('recalc', 'cli.cmd.recalc').actionCtx(async (options, _args, ctx) => {
    const file = input(options);
    if (file === undefined) {
        ctx.fail('open-excel.cli.file-required', { message: 'a project file is required' });
        return 2;
    }
    const opened = service.open(file);
    try {
        const report = service.recalc(opened.handle);
        ctx.emitResult(report, describeRecalc(report));
        return 0;
    } finally {
        service.close(opened.handle);
    }
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
        const opened = service.open(file);
        try {
            const report = service.exportWorkbook(opened.handle, format as ExportFormat, out, false);
            ctx.emitResult(report, `exported ${out} format=${report.format} losses=${report.losses.length}`);
            return 0;
        } finally {
            service.close(opened.handle);
        }
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
