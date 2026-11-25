/**
 * Error contract checks against a stub binding.
 *
 * The stub replaces the native display text on purpose: the SDK must still
 * branch on the machine-readable `code` and typed `params`, never on prose.
 * Unknown codes, malformed payloads and the `binding/failure` fallback all run
 * through the same path.
 */

import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

import { isOpenExcelError, parseErrorParams, setErrorLocale, type OpenExcelError } from '../src/errors.js';
import * as service from '../src/service.js';

const dir = mkdtempSync(join(tmpdir(), 'open-excel-stub-'));
const stubPath = join(dir, 'stub-binding.cjs');
writeFileSync(
    stubPath,
    [
        'function failure() {',
        '    const next = JSON.parse(process.env.OPEN_EXCEL_STUB || "{}");',
        '    const error = new Error(next.message || "the native text was replaced");',
        '    if (next.code !== undefined) { error.code = next.code; }',
        '    if (next.detail !== undefined) { error.detail = next.detail; }',
        '    return error;',
        '}',
        'module.exports = {',
        '    version() { return "0.0.0-stub"; },',
        '    workbook_create() { throw failure(); },',
        '};',
    ].join('\n'),
);
process.env.OPEN_EXCEL_NATIVE = stubPath;

function capture(): OpenExcelError {
    try {
        service.create({ title: 'stub' });
    } catch (error) {
        assert.ok(isOpenExcelError(error), 'every service failure is an OpenExcelError');
        return error;
    }
    return assert.fail('the stub always throws') as never;
}

test('structured binding failures keep their code and typed params', () => {
    process.env.OPEN_EXCEL_STUB = JSON.stringify({
        message: 'REPLACED NATIVE TEXT',
        code: 'object/not-found',
        detail: { params: { kind: 'sheet', id: 'Sheet9' } },
    });
    const error = capture();
    assert.equal(error.code, 'object/not-found');
    assert.deepEqual(error.params, { kind: 'sheet', id: 'Sheet9' });
    assert.equal(error.message, 'the `sheet` `Sheet9` was not found', 'the product copy replaces the native text');
});

test('strict params reject missing, extra and mistyped keys', () => {
    assert.equal(parseErrorParams('io/read', { path: 'book.xlsx' }).ok, false, 'a missing OS kind is refused');
    assert.equal(
        parseErrorParams('io/exists', { path: 'book.xlsx', legacy: true }).ok,
        false,
        'a key outside the frozen shape is refused',
    );
    assert.equal(
        parseErrorParams('document/revision-conflict', { actual: '7' }).ok,
        false,
        'a revision must be a number',
    );
    assert.equal(parseErrorParams('weird/code', {}).ok, false, 'an unknown code has no frozen shape');
    assert.deepEqual(
        parseErrorParams('object/not-found', { kind: 'sheet', id: 'Sheet1' }),
        { ok: true, params: { kind: 'sheet', id: 'Sheet1' } },
    );
});

test('a malformed payload falls back to binding/failure with the original code and text', () => {
    process.env.OPEN_EXCEL_STUB = JSON.stringify({ message: 'disk said no', code: 'io/exists', detail: { params: 'not-an-object' } });
    const error = capture();
    assert.equal(error.code, 'binding/failure');
    assert.deepEqual(error.params, { reason: 'disk said no', code: 'io/exists' });
    assert.equal(error.message, 'the native binding reported a failure: disk said no', 'the binding text is kept as payload');
});

test('missing required params land on binding/failure', () => {
    process.env.OPEN_EXCEL_STUB = JSON.stringify({ message: 'could not read', code: 'io/read', detail: { params: { path: 'book.xlsx' } } });
    const error = capture();
    assert.equal(error.code, 'binding/failure');
    assert.deepEqual(error.params, { reason: 'could not read', code: 'io/read' });
});

test('prose that looks like a legacy code is not parsed', () => {
    process.env.OPEN_EXCEL_STUB = JSON.stringify({ message: 'workbook/not-found: old binding text' });
    const error = capture();
    assert.equal(error.code, 'binding/failure', 'a message never decides the code');
    assert.deepEqual(error.params, { reason: 'workbook/not-found: old binding text' });
});

test('unknown codes land on binding/failure and keep the original code in params', () => {
    process.env.OPEN_EXCEL_STUB = JSON.stringify({ message: 'native text kept', code: 'weird/code' });
    const error = capture();
    assert.equal(error.code, 'binding/failure');
    assert.deepEqual(error.params, { reason: 'native text kept', code: 'weird/code' });
});

test('a failure without any code surfaces as binding/failure', () => {
    process.env.OPEN_EXCEL_STUB = JSON.stringify({ message: 'boom' });
    const error = capture();
    assert.equal(error.code, 'binding/failure');
    assert.deepEqual(error.params, { reason: 'boom' });
    assert.equal(error.message, 'the native binding reported a failure: boom');
});

test('zh-CN rendering changes only the copy, never the machine fields', () => {
    setErrorLocale('zh-CN');
    try {
        process.env.OPEN_EXCEL_STUB = JSON.stringify({
            message: 'REPLACED NATIVE TEXT',
            code: 'object/not-found',
            detail: { params: { kind: 'sheet', id: 'Sheet9' } },
        });
        const error = capture();
        assert.equal(error.code, 'object/not-found');
        assert.deepEqual(error.params, { kind: 'sheet', id: 'Sheet9' });
        assert.equal(error.message, '未找到 `sheet` `Sheet9`');
    } finally {
        setErrorLocale('en');
    }
});

test('switching locale never moves the code or the params', () => {
    process.env.OPEN_EXCEL_STUB = JSON.stringify({
        message: 'REPLACED NATIVE TEXT',
        code: 'document/revision-conflict',
        detail: { params: { actual: 7, expected: 6 } },
    });
    const en = capture();
    setErrorLocale('zh-CN');
    try {
        const zh = capture();
        assert.equal(zh.code, en.code);
        assert.deepEqual(zh.params, en.params);
        assert.notEqual(zh.message, en.message, 'only the copy changes');
    } finally {
        setErrorLocale('en');
    }
});
