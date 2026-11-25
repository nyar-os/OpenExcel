/**
 * CLI locale parity for the installed `en-US` and `zh-CN` catalogs.
 *
 * The same failure under both locales must produce the same machine-readable
 * envelope — stable `code`, identical interpolated params and the same exit
 * code — while only the human copy changes. An unknown binding code lands on
 * the structured `binding/failure` fallback in every locale.
 */

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';

const cliPath = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..', 'dist', 'cli.js');

interface Envelope {
    readonly schema: string;
    readonly ok: boolean;
    readonly error?: { readonly code: string; readonly message: string };
}

interface Run {
    readonly exit: number | null;
    readonly envelope: Envelope;
}

const baseEnv = { ...process.env, LOCALE: '', LANG: '', LC_ALL: '', NO_COLOR: '1', OPEN_EXCEL_NATIVE: '' };

function run(args: readonly string[], extraEnv: Readonly<Record<string, string>> = {}): Run {
    const result = spawnSync(process.execPath, [cliPath, ...args], {
        encoding: 'utf8',
        env: { ...baseEnv, ...extraEnv },
    });
    assert.equal(result.error, undefined, 'the CLI must spawn');
    return { exit: result.status, envelope: JSON.parse(result.stdout) as Envelope };
}

test('a cli-level failure keeps code and exit and differs only in copy', () => {
    const en = run(['--locale', 'en-US', 'summary', '--json']);
    const zh = run(['--locale', 'zh-CN', 'summary', '--json']);
    assert.equal(zh.exit, en.exit, 'the exit code is locale-independent');
    assert.equal(en.envelope.error?.code, 'open-excel.cli.file-required');
    assert.equal(zh.envelope.error?.code, en.envelope.error?.code, 'the machine code is locale-independent');
    assert.equal(en.envelope.error?.message, 'a project file is required');
    assert.equal(zh.envelope.error?.message, '需要项目文件');
});

test('structured params interpolate identically in both locales', () => {
    const args = ['row', '--insert', '1048576', '--count', '2', 'book.xlsx', '--json'];
    const en = run(['--locale', 'en-US', ...args]);
    const zh = run(['--locale', 'zh-CN', ...args]);
    assert.equal(zh.exit, en.exit);
    assert.equal(en.envelope.error?.code, 'open-excel.cli.row-band-invalid');
    assert.equal(zh.envelope.error?.code, en.envelope.error?.code);
    assert.match(en.envelope.error?.message ?? '', /1048576\.\.1048577/, 'the en copy carries the params');
    assert.match(zh.envelope.error?.message ?? '', /1048576\.\.1048577/, 'the zh copy carries the same params');
    assert.notEqual(en.envelope.error?.message, zh.envelope.error?.message, 'only the human copy changes');
});

const dir = mkdtempSync(path.join(tmpdir(), 'open-excel-cli-stub-'));
const stubPath = path.join(dir, 'stub-binding.cjs');
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

function sdkRun(locale: string, stub: unknown): Run {
    return run(['--locale', locale, 'create', '--out', 'book.xlsx', '--json'], {
        OPEN_EXCEL_NATIVE: stubPath,
        OPEN_EXCEL_STUB: JSON.stringify(stub),
    });
}

test('an sdk failure reaches the envelope under its own code in both locales', () => {
    const failure = {
        message: 'REPLACED NATIVE TEXT',
        code: 'object/not-found',
        detail: { params: { kind: 'sheet', id: 'Sheet9' } },
    };
    const en = sdkRun('en-US', failure);
    const zh = sdkRun('zh-CN', failure);
    assert.equal(zh.exit, en.exit);
    assert.equal(en.envelope.error?.code, 'object/not-found', 'the SDK code replaces the framework fallback');
    assert.equal(zh.envelope.error?.code, en.envelope.error?.code);
    assert.equal(en.envelope.error?.message, 'the `sheet` `Sheet9` was not found');
    assert.equal(zh.envelope.error?.message, '未找到 `sheet` `Sheet9`', 'the session locale renders the copy');
});

test('an unknown code falls back to binding/failure in every locale', () => {
    const failure = { message: 'native text kept', code: 'weird/code' };
    const en = sdkRun('en-US', failure);
    const zh = sdkRun('zh-CN', failure);
    assert.equal(zh.exit, en.exit);
    assert.equal(en.envelope.error?.code, 'binding/failure', 'an unknown code is structured, not guessed');
    assert.equal(zh.envelope.error?.code, en.envelope.error?.code);
    assert.equal(en.envelope.error?.message, 'the native binding reported a failure: native text kept');
    assert.equal(zh.envelope.error?.message, '原生绑定返回了失败：native text kept');
});

test('help renders the registered locale catalog', () => {
    const zh = spawnSync(process.execPath, [cliPath, '--locale', 'zh-CN', '--help'], { encoding: 'utf8', env: baseEnv });
    assert.equal(zh.status, 0);
    assert.match(zh.stdout, /打开项目文件并打印其工作表与单元格/, 'zh-CN command copy ships with the CLI');
    const en = spawnSync(process.execPath, [cliPath, '--locale', 'en-US', '--help'], { encoding: 'utf8', env: baseEnv });
    assert.equal(en.status, 0);
    assert.match(en.stdout, /Open a project file and print its sheets and cells/);
});
