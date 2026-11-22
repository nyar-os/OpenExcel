# @wae/open-excel-skills

Agent skill documents for OpenExcel. The package ships Markdown only: no runtime library, ZIP/XML code or workbook algorithms.

## Usage

Install it next to `@wae/open-excel`, then ask your agent to use the `open-excel` skill:

```text
Use the open-excel skill to open q1-results.xlsx, add a Total column with
=SUM(B2:D2), and export a PDF for review.
```

The skill drives the `open-excel` CLI and starts with `open-excel capabilities --json`.

## Related

- [Product overview](../../../readme.md)
- [`@wae/open-excel`](../open-excel/README.md)
- [Agent reference](../../../documentation/agents.md)
