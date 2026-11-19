# open-excel-ui

VMZ static UI for the OpenExcel Studio window. `vmz build . --profile static`
emits `dist/static`, which the repository staging step copies to
`projects/packages/open-excel/dist/gui` for the native WebView host.

```bash
pnpm --filter @wae/open-excel-ui run check   # vmz check
pnpm --filter @wae/open-excel-ui run build   # vmz build --profile static
pnpm --filter @wae/open-excel-ui run serve   # local preview
```

## Status

Shell only. The page talks to the host over `chrome.webview.postMessage`
JSON-RPC (`ready`, `capabilities`, `workbook.create`, `workbook.open`,
`workbook.save`, `workbook.inspect`, `workbook.export`, `workbook.close`,
`close`). No workbook is drawn and no cell grid is synthesized client-side: with
no native host the page shows an empty state and a "Preview mode" notice, because
the spreadsheet engine is not implemented.
