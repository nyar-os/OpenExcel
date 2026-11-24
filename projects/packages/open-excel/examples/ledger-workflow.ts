/**
 * Ledger calculation workflow, the runnable form of `documentation/agents.md`.
 *
 * The script creates a workbook, writes numeric seeds and explicit
 * `{ formula }` cells in one atomic batch, proves the batch evaluated on
 * commit (an idempotent `recalc` changes nothing), asserts the computed
 * results `D2=260` and `D3=200`, inspects the tracked workbook, exports the
 * XLSX package and saves the native project, then reopens both files and
 * checks the formula source and the cached result in each. It exits non-zero
 * on the first failed assertion, so CI runs the documented workflow verbatim
 * instead of trusting the prose.
 *
 * Run from the repository root with a built binding installed or staged:
 *
 * ```bash
 * pnpm --filter @wae/open-excel run build:test
 * node projects/packages/open-excel/dist-test/examples/ledger-workflow.js
 * ```
 *
 * On hosts without an installed platform package point `OPEN_EXCEL_NATIVE` at
 * a freshly built addon renamed to `.node` (`open_excel_napi.node`).
 */

import assert from 'node:assert/strict';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { OpenExcel } from '../src/index.js';
import { nativeBinding } from '../src/native.js';

async function main(): Promise<void> {
    if (nativeBinding() === undefined) {
        throw new Error('this example needs a built native binding; install the platform package or set OPEN_EXCEL_NATIVE');
    }
    const dir = mkdtempSync(join(tmpdir(), 'open-excel-ledger-'));
    const xlsxPath = join(dir, 'ledger-final.xlsx');
    const projectPath = join(dir, 'ledger-final.oexcel');

    // 1. One atomic batch: numeric seeds plus explicit formula cells.
    const excel = new OpenExcel({ locale: 'en' });
    const book = await excel.create({ title: 'Ledger' });
    const edited = await excel.edit(book, {
        expectedRevision: book.revision,
        operations: [
            {
                op: 'range.set',
                sheet: 'Sheet1',
                start: 'A1',
                rows: [
                    ['Region', 'Q1', 'Q2', 'Total'],
                    ['North', 120, 140, null],
                    ['South', 90, 110, null],
                ],
            },
            { op: 'cell.set', sheet: 'Sheet1', reference: 'D2', formula: '=B2+C2' },
            { op: 'cell.set', sheet: 'Sheet1', reference: 'D3', formula: '=B3+C3' },
            { op: 'format.set', sheet: 'Sheet1', range: 'A1:D1', style: { bold: true, fill: '#eef2ff' } },
        ],
    });
    assert.equal(edited.workbook.revision, book.revision + 1, 'the batch advances the revision once');

    // 2. `edit` evaluates on commit, so a following `recalc` is idempotent.
    const recalc = await excel.recalc(edited.workbook);
    assert.equal(recalc.changed, 0, 'the committed formulas already carry their results');
    const north = await excel.cell(edited.workbook, { sheet: 'Sheet1', address: 'D2' });
    const south = await excel.cell(edited.workbook, { sheet: 'Sheet1', address: 'D3' });
    assert.equal(north.value, 260);
    assert.equal(south.value, 200);

    // 3. Inspect the tracked workbook; error findings fail the delivery.
    const inspection = await excel.inspect(edited.workbook, { checks: ['formulas', 'references', 'cycles'] });
    assert.equal(inspection.passed, true, JSON.stringify(inspection));

    // 4. Write the exchange package and the native project.
    const xlsx = await excel.export(edited.workbook, { format: 'xlsx', target: xlsxPath, overwrite: true });
    assert.equal(xlsx.target, xlsxPath);
    await excel.save(edited.workbook, { target: projectPath, overwrite: true });
    await excel.dispose();

    // 5. Reopen both artifacts: formula source and cached result in each.
    const reopened = new OpenExcel();
    for (const path of [projectPath, xlsxPath]) {
        const restored = await reopened.open({ path });
        const total = await reopened.cell(restored, { sheet: 'Sheet1', address: 'D2' });
        assert.equal(total.formula, '=B2+C2', `${path} keeps the formula source`);
        assert.equal(total.value, 260, `${path} keeps the computed result`);
        const other = await reopened.cell(restored, { sheet: 'Sheet1', address: 'D3' });
        assert.equal(other.value, 200);
    }
    await reopened.dispose();

    console.log(`ledger workflow ok: D2=260 D3=200, revision ${edited.workbook.revision}, reopened ${projectPath} and ${xlsxPath}`);
}

await main();
