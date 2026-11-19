# @wae/open-excel-skills

Agent skills for OpenExcel. The package ships skill documents only; it is not a
runtime library and contains no ZIP, XML, formula or drawing algorithms.

## Status

Scaffold only. `skills/open-excel/SKILL.md` records the routing and evidence
rules for create, edit, template fill, review and delivery, but no SDK or CLI
capability exists yet, so none of those routes is executable.

## Layout

```text
packages/open-excel-skills/
├── package.json
├── README.md
└── skills/
    └── open-excel/
        └── SKILL.md
```

Design rules, the format/support matrix and executable examples are added with
the first implemented release. The skill must never be shipped claiming methods,
commands or flags that the installed SDK/CLI does not expose.
