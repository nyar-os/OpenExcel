# @wae/open-excel

TypeScript surface for OpenExcel: one declared workbook model behind the command
line, the studio window and agent workflows. The session entry point is
`new OpenExcel(options?)`.

## Status

Declared surface only. `new OpenExcel(options?)` creates a session synchronously,
and the package exports the frozen types: workbook summaries with sheets and
cells, inspect/save/export reporting, the stable error codes, and the workbook
format registry (`.xlsx`/`.xlsm`/`.xltx`/`.xls`/`.csv`/`.tsv`/`.ods`/`.json`/
`.xml`/`.pdf`).

Every capability still rejects with `capability/not-implemented`. Classification
is not a support claim: reads, writes, inspection and export are reported by
`capabilities()` once the shared workbook engine lands, never inferred from an
extension or a method name.

## Boundaries

- No workbook parsing, formula evaluation, pivot aggregation or ZIP/XML code in
  this package.
- No second workbook model: the CLI, the studio and agents read the same
  capability matrix.
- Native access goes through the Node-API binding in `projects/crates/napi`.

## Scripts

```bash
pnpm --filter @wae/open-excel run typecheck
pnpm --filter @wae/open-excel run build
pnpm --filter @wae/open-excel run test
```
