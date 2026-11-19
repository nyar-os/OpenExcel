/**
 * Smoke tests for the declared workbook surface.
 *
 * They assert shell semantics only: one stable error code for unimplemented
 * capabilities, a session that closes idempotently, synchronous argument
 * validation, and a format registry that classifies the whole family. No
 * workbook, renderer or import behaviour is claimed here.
 */

import assert from 'node:assert/strict';
import test from 'node:test';

import { formatFromPath, formatInfo, OpenExcel, OpenExcelError, workbookFormats } from './index.js';
import type { WorkbookInput } from './index.js';
import { nativeBinding } from './native.js';

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
