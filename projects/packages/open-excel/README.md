# @wae/open-excel

The session API behind the `open-excel` CLI: one workbook handle for reading, editing, inspecting, searching, recalculating and exporting.

## API surface

`new OpenExcel(options?)` creates a session. Nothing is opened until an async method runs; `dispose()` releases every workbook handle the session still tracks.

```ts
const excel = new OpenExcel({ locale: 'en' });

await excel.capabilities();                       // formats, checks, formulas, renderers, export targets

const book = await excel.create({ title: 'Ledger' });
const opened = await excel.open({ path: 'ledger.oexcel' });
await excel.summary(book);                        // sheets, revisions and cell counts
await excel.save(book, { target: 'copy.oexcel', overwrite: true });

await excel.edit(book, { operations });           // atomic batch; one revision advance
await excel.cell(book, { sheet: 'Sheet1', address: 'B4' });
await excel.range(book, { sheet: 'Sheet1', range: 'A1:D10' });
await excel.recalc(book);                         // recompute every formula
await excel.find(book, { query: 'invoice' });     // values and formulas, 500-hit cap

await excel.export(book, { format: 'xlsx', target: 'ledger.xlsx', overwrite: true });
const svg = await excel.render(book, { sheet: 'Sheet1' });
await excel.dispose();
```

`edit()` accepts these operations: `cell.set`, `range.set`, `range.copy`, `range.move`, `range.swap`, `range.sort`, `range.fillSeries`, `range.transpose`, `range.clear`, `format.set`, `merge.set`, `sheet.freeze`, `sheet.setVisibility`, `sheet.add`, `sheet.rename`, `sheet.remove`, `sheet.duplicate`, `sheet.move`, `row.insert`, `row.delete`, `column.insert`, `column.delete`, `row.setHeight`, `column.setWidth` and `cell.replaceText`. Pass `expectedRevision` to reject a stale writer.

Export targets are `xlsx`, `csv`, `tsv`, `json`, `html` and `pdf`; `xlsm` export is not available. CSV, TSV and HTML write the primary sheet and report skipped sheets as losses. PDF is partial and reports its own loss codes. `open()` reads a path to `.oexcel`, `.xlsx`, `.xlsm`, `.csv` or `.tsv`; in-memory bytes and non-edit intents are rejected. Failures are `OpenExcelError` instances with one of the stable codes from `errorCodes`.

## Example

```ts
import { OpenExcel } from '@wae/open-excel';

const excel = new OpenExcel();
const book = await excel.create({ title: 'Ledger' });
await excel.edit(book, {
    operations: [{ op: 'cell.set', sheet: 'Sheet1', reference: 'A1', formula: '=SUM(B1:B3)' }],
});
console.log((await excel.cell(book, { address: 'A1' })).value);
```

## Related packages

- [OpenExcel root README](../../../readme.md)
- [@wae/open-excel-skills](../open-excel-skills/README.md)
- [@wae/open-excel-ui](../open-excel-ui/README.md)
