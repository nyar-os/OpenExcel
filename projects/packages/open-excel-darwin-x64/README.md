# @wae/open-excel-darwin-x64

Prebuilt macOS x64 native addon for [`@wae/open-excel`](../open-excel/README.md). **Not a public import target.** It is installed automatically as an optional dependency on matching `os` and `cpu`; application code imports `@wae/open-excel`.

## Example

```ts
import { OpenExcel } from '@wae/open-excel';

const excel = new OpenExcel();
const book = await excel.create({ title: 'Ledger' });
await excel.cell(book, { sheet: 'Sheet1', address: 'A1' });
```
