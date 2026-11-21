# @wae/open-excel-linux-arm64

Linux arm64 platform package for the OpenExcel native addon.

This package carries the prebuilt `open-excel.linux-arm64.node` file produced by
the `stage:native` step after `wae native build`. The repository ships no compiled
binary; the file appears in a staged or published artifact only.

The addon maps values only and contains no workbook logic. Until the spreadsheet
engine lands, every operation reports `capability/not-implemented`.
