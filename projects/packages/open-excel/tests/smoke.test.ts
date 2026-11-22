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
