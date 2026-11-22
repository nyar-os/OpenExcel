# @wae/open-excel-ui

Desktop shell for OpenExcel Studio. This private package contains the VMZ static UI that is staged into the `@wae/open-excel` GUI asset; it is not published on its own.

## Usage

```bash
open-excel gui
```

The window talks to the native host over `chrome.webview.postMessage` JSON-RPC (`ready`, `capabilities`, `workbook.create`, `workbook.open`, `workbook.save`, `workbook.inspect`, `workbook.export`, `workbook.close`, `close`). Without a WebView host it shows an empty state and a preview notice; no cell grid is synthesized client-side.
