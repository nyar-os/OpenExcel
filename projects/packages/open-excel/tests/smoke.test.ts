/**
 * Smoke tests for the declared workbook surface.
 *
 * They assert shell semantics only: one stable error code for unimplemented
 * capabilities, a session that closes idempotently, synchronous argument
 * validation, a format registry that classifies the whole family, and one
 * structural edit that needs the native engine. No renderer or import behaviour
 * is claimed here.
 */

import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';

import { formatFromPath, formatInfo, OpenExcel, OpenExcelError, workbookFormats } from '../src/index.js';
import type { Workbook, WorkbookInput } from '../src/index.js';
import { nativeBinding } from '../src/native.js';

const installed = (() => {
    try {
        return nativeBinding() !== undefined;
    } catch {
        return false;
    }
})();
const withAddon = installed ? 'a built addon is installed' : false;

test('session construction is synchronous and side-effect free', async () => {
    const excel = new OpenExcel();
    assert.equal(excel.locale, 'en');
    assert.equal(excel.disposed, false);
    await excel.dispose();
    assert.equal(excel.disposed, true);
});

test('without an installed addon engine calls keep capability/not-implemented', { skip: withAddon }, async () => {
    const excel = new OpenExcel({ locale: 'zh-CN' });
    assert.equal(excel.locale, 'zh-CN');
    const calls: (() => Promise<unknown>)[] = [
        () => excel.capabilities(),
        () => excel.create({ title: 'Ledger', template: 'blank' }),
        () => excel.open({ path: 'ledger.oexcel' }),
        () => excel.inspect({ path: 'ledger.xlsx' }),
    ];
    for (const call of calls) {
        await assert.rejects(call, (error: unknown) => {
            assert.ok(error instanceof OpenExcelError);
            assert.equal(error.code, 'capability/not-implemented');
            assert.ok(error.details.feature);
            return true;
        });
    }
    await excel.dispose();
});

test('dispose is idempotent and closes the session', async () => {
    const excel = new OpenExcel();
    await excel.dispose();
    await excel.dispose();
    await assert.rejects(
        excel.capabilities(),
        (error: unknown) => error instanceof OpenExcelError && error.code === 'session/disposed',
    );
});

test('dispose resolves after a rejected create and open', { skip: withAddon }, async () => {
    const excel = new OpenExcel();
    await assert.rejects(excel.create({ title: 'Ledger' }));
    await assert.rejects(excel.open({ path: 'ledger.oexcel' }));
    await excel.dispose();
    await excel.dispose();
    assert.equal(excel.disposed, true);
});

test('open rejects ambiguous input before reaching the engine', async () => {
    const excel = new OpenExcel();
    const inputs = [
        {} as unknown as WorkbookInput,
        { path: 'a.oexcel', bytes: new Uint8Array() } as unknown as WorkbookInput,
        { bytes: 'not bytes' } as unknown as WorkbookInput,
        { path: '' } as unknown as WorkbookInput,
    ];
    for (const input of inputs) {
        await assert.rejects(
            excel.open(input),
            (error: unknown) => error instanceof OpenExcelError && error.code === 'session/argument-invalid',
        );
    }
    await excel.dispose();
});

test('create rejects empty titles synchronously', async () => {
    const excel = new OpenExcel();
    await assert.rejects(
        excel.create({ title: '' }),
        (error: unknown) => error instanceof OpenExcelError && error.code === 'session/argument-invalid',
    );
    await assert.rejects(
        excel.create({ template: '' }),
        (error: unknown) => error instanceof OpenExcelError && error.code === 'session/argument-invalid',
    );
    await excel.dispose();
});

test('save, edit, export and render validate arguments before reaching the engine', async () => {
    const excel = new OpenExcel();
    const workbook = { id: 'workbook-0', revision: 0, title: 'Fixture', sheets: [], cells: 0 } as unknown as Workbook;
    const invalid = (error: unknown) => error instanceof OpenExcelError && error.code === 'session/argument-invalid';

    await assert.rejects(excel.save(workbook, { target: '' }), invalid);
    await assert.rejects(excel.edit(workbook, { operations: [] }), invalid);
    await assert.rejects(
        excel.edit(workbook, { expectedRevision: -1, operations: [{ op: 'sheet.add', name: 'Second' }] }),
        invalid,
    );
    await assert.rejects(
        excel.edit(workbook, { operations: [{} as never] }),
        invalid,
    );
    await assert.rejects(excel.export(workbook, { format: 'csv', target: '' }), invalid);
    await assert.rejects(excel.cell(workbook, { address: ' ' }), invalid);
    await assert.rejects(excel.range(workbook, { range: '' }), invalid);
    await assert.rejects(
        excel.render(workbook, { format: 'png' }),
        (error: unknown) => error instanceof OpenExcelError && error.code === 'format/unsupported',
    );
    await excel.dispose();
});

test('tracked reads reject unknown workbooks before reaching the engine', async () => {
    const excel = new OpenExcel();
    const workbook = { id: 'workbook-404', revision: 0, title: 'Fixture', sheets: [], cells: 0 } as unknown as Workbook;
    const missing = (error: unknown) => error instanceof OpenExcelError && error.code === 'workbook/not-found';
    await assert.rejects(excel.cell(workbook, { address: 'A1' }), missing);
    await assert.rejects(excel.range(workbook, { range: 'A1:B2' }), missing);
    await assert.rejects(excel.recalc(workbook), missing);
    await excel.dispose();
});

test('tracked operations reject unknown workbooks before reaching the engine', async () => {
    const excel = new OpenExcel();
    const workbook = { id: 'workbook-404', revision: 0, title: 'Fixture', sheets: [], cells: 0 } as unknown as Workbook;
    const missing = (error: unknown) => error instanceof OpenExcelError && error.code === 'workbook/not-found';
    await assert.rejects(excel.summary(workbook), missing);
    await assert.rejects(excel.save(workbook, { target: 'fixture.oexcel' }), missing);
    await assert.rejects(excel.export(workbook, { format: 'csv', target: 'fixture.csv' }), missing);
    await excel.dispose();
});

test('format registry classifies the workbook format family', () => {
    const expected = ['xlsx', 'xlsm', 'xltx', 'xls', 'csv', 'tsv', 'ods', 'json', 'xml', 'pdf'];
    assert.deepEqual([...workbookFormats].sort(), [...expected].sort());

    assert.equal(formatInfo('xlsx').defaultIntent, 'edit');
    assert.equal(formatInfo('xltx').defaultIntent, 'template');
    assert.equal(formatInfo('csv').defaultIntent, 'data');
    assert.equal(formatInfo('csv').opensAs, 'data');
    assert.equal(formatInfo('xlsm').macroEnabled, true);
    assert.equal(formatInfo('xltx').family, 'template');
    assert.equal(formatInfo('xls').family, 'legacy');
    assert.equal(formatInfo('ods').markup, 'odf');
    assert.equal(formatInfo('pdf').opensAs, 'none');
    assert.equal(formatInfo('pdf').defaultIntent, null);

    assert.equal(formatFromPath('Ledger.XLSX'), 'xlsx');
    assert.equal(formatFromPath('sheet.numbers'), undefined);
    assert.equal(formatFromPath('no-extension'), undefined);
});

test('a native addon can insert a row and move its references', { skip: installed ? false : 'a built addon is not installed' }, async (context) => {
    const excel = new OpenExcel();
    const created = await excel.create({ title: 'Structure' });
    const seeded = await excel.edit(created, {
        operations: [
            { op: 'cell.set', sheet: 'Sheet1', reference: 'A1', value: 1 },
            { op: 'cell.set', sheet: 'Sheet1', reference: 'A2', value: 2 },
            { op: 'cell.set', sheet: 'Sheet1', reference: 'B2', formula: '=A2' },
        ],
    });
    let edited;
    try {
        edited = await excel.edit(seeded.workbook, { operations: [{ op: 'row.insert', sheet: 'Sheet1', row: 2 }] });
    } catch (error) {
        if (error instanceof OpenExcelError && error.code === 'object/unsupported' && error.message.includes('row.insert')) {
            context.skip('the installed addon predates `row.insert`');
            await excel.dispose();
            return;
        }
        throw error;
    }
    assert.equal(edited.revision, 2);
    const moved = await excel.cell(edited.workbook, { sheet: 'Sheet1', address: 'A3' });
    assert.equal(moved.value, 2);
    const formula = await excel.cell(edited.workbook, { sheet: 'Sheet1', address: 'B3' });
    assert.equal(formula.formula, '=A3');
    await excel.dispose();
});

test('cross-sheet moves, swaps, names and merges carry formula sources and fail atomically', { skip: installed ? false : 'a built addon is not installed' }, async () => {
    const excel = new OpenExcel();
    const created = await excel.create({ title: 'Bindings' });
    const seeded = await excel.edit(created, {
        operations: [
            { op: 'sheet.add', name: 'Data' },
            { op: 'range.set', sheet: 'Sheet1', start: 'A1', rows: [['name', 'total'], ['alpha', { formula: '=1+2' }]] },
        ],
    });

    // A cross-sheet move carries the formula source and its computed result.
    const moved = await excel.edit(seeded.workbook, {
        operations: [{ op: 'range.move', sheet: 'Sheet1', range: 'A1:B2', target: 'D4', targetSheet: 'Data' }],
    });
    assert.equal(moved.revision, 2);
    const movedTotal = await excel.cell(moved.workbook, { sheet: 'Data', address: 'E5' });
    assert.equal(movedTotal.formula, '=1+2');
    assert.equal(movedTotal.value, 3);
    const cleared = await excel.cell(moved.workbook, { sheet: 'Sheet1', address: 'B2' });
    assert.equal(cleared.value, null);
    assert.equal(cleared.formula, null);

    // A workbook name records the moved block on its new sheet.
    const named = await excel.edit(moved.workbook, {
        operations: [{ op: 'name.define', name: 'Totals', sheet: 'Data', range: 'D4:E5' }],
    });
    assert.deepEqual(named.workbook.names, [{ name: 'Totals', sheet: 'Data', range: 'D4:E5' }]);

    // A cross-sheet swap exchanges both rectangles verbatim.
    const swapped = await excel.edit(named.workbook, {
        operations: [
            { op: 'range.set', sheet: 'Sheet1', start: 'A1', rows: [['x', 10], ['y', 20]] },
            { op: 'range.swap', sheet: 'Data', range: 'D4:E5', target: 'A1', targetSheet: 'Sheet1' },
        ],
    });
    const swappedTotal = await excel.cell(swapped.workbook, { sheet: 'Sheet1', address: 'B2' });
    assert.equal(swappedTotal.formula, '=1+2');
    assert.equal(swappedTotal.value, 3);
    const swappedName = await excel.cell(swapped.workbook, { sheet: 'Data', address: 'D4' });
    assert.equal(swappedName.value, 'x');

    // A merge blocks a move that would write over it, and the failing batch
    // leaves the revision and the names untouched.
    const merged = await excel.edit(swapped.workbook, {
        operations: [{ op: 'merge.set', sheet: 'Data', start: 'G1', end: 'H1', merge: true }],
    });
    const revisionBefore = merged.revision;
    await assert.rejects(
        excel.edit(merged.workbook, {
            operations: [
                { op: 'name.define', name: 'Blocked', sheet: 'Data', range: 'G1:H1' },
                { op: 'range.move', sheet: 'Sheet1', range: 'A1:B2', target: 'G1', targetSheet: 'Data' },
            ],
        }),
        (error) => error instanceof OpenExcelError && error.code === 'session/argument-invalid',
    );
    const after = await excel.summary(merged.workbook);
    assert.equal(after.revision, revisionBefore);
    assert.equal(after.names.filter((entry) => entry.name === 'Blocked').length, 0);
    await excel.dispose();
});

test('real binding failures carry the typed params of their code', { skip: installed ? false : 'a built addon is not installed' }, async () => {
    const excel = new OpenExcel();
    try {
        await assert.rejects(
            excel.open({ path: join(tmpdir(), 'open-excel-missing-file.oexcel') }),
            (error: unknown) => {
                assert.ok(error instanceof OpenExcelError);
                const failure = error.failure;
                assert.equal(failure.code, 'io/read');
                if (failure.code !== 'io/read') return false;
                assert.equal(failure.params.kind, 'NotFound');
                assert.equal(typeof failure.params.path, 'string');
                return true;
            },
        );

        const created = await excel.create({ title: 'Params' });
        await assert.rejects(
            excel.cell(created, { sheet: 'Nope', address: 'A1' }),
            (error: unknown) => {
                assert.ok(error instanceof OpenExcelError);
                const failure = error.failure;
                assert.equal(failure.code, 'object/not-found');
                if (failure.code !== 'object/not-found') return false;
                assert.equal(failure.params.kind, 'sheet');
                assert.equal(failure.params.id, 'Nope');
                return true;
            },
        );

        const dir = mkdtempSync(join(tmpdir(), 'open-excel-params-'));
        const target = join(dir, 'ledger.oexcel');
        writeFileSync(target, 'occupied\n');
        await assert.rejects(
            excel.save(created, { target }),
            (error: unknown) => {
                assert.ok(error instanceof OpenExcelError);
                const failure = error.failure;
                assert.equal(failure.code, 'io/exists');
                if (failure.code !== 'io/exists') return false;
                assert.equal(failure.params.path, target);
                return true;
            },
        );
    } finally {
        await excel.dispose();
    }
});
