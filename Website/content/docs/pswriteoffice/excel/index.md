---
title: "Automate Excel Workbooks"
description: "Build, inspect, validate, compare, repair, and publish workbook workflows from PowerShell."
layout: docs
---

Excel is the largest PSWriteOffice family with 155 exported commands. It covers workbook creation and reading, sheet and range operations, formulas, styling, tables, charts, pivots, validation, comments, images, links, templates, dashboards, protection, accessibility, comparison, repair, and streaming contracts.

## Create a workbook from data

Use `New-OfficeExcel` with `Add-OfficeExcelSheet`, then add tables, formulas, charts, and report components inside each sheet context. The report DSL includes titles, paragraphs, sections, callouts, KPI rows, tables, legends, spacers, and dashboard charts for repeatable operational output.

```powershell
$records = @(
    [pscustomobject]@{ Region = 'EMEA'; Revenue = 98000 }
    [pscustomobject]@{ Region = 'APAC'; Revenue = 143000 }
)

New-OfficeExcel -Path '.\Output\Revenue.xlsx' {
    Add-OfficeExcelSheet -Name 'Sales' {
        Add-OfficeExcelTable -InputObject $records -AutoFilter -AutoFit
        Add-OfficeExcelChart -Type ColumnClustered -Title 'Revenue by region'
    }
}
```

## Work with existing workbooks

The read surface can return used ranges, tables, named ranges, formulas, comments, validation, rich text, worksheet views, page breaks, pivots, summaries, preflight data, and streaming capabilities. Targeted commands update cells, rows, columns, styles, formulas, links, page setup, print settings, themes, worksheet visibility, active sheet, filters, and write reservations.

## Validate before delivery

- `Get-OfficeExcelPreflight` and `Get-OfficeExcelRuntimePreflight` report readiness before an operation.
- `Test-OfficeExcelWorkbook` checks workbook integrity.
- `Test-OfficeExcelAccessibility` supports accessible-delivery gates.
- `Compare-OfficeExcelWorkbook` and `Compare-OfficeExcelRange` make change evidence explicit.
- `Repair-OfficeExcelWorkbook` is a deliberate repair path, not an implicit side effect of reading.

Templates, joins, merges, sheet ordering, workbook copying, HTML review, and delimited import/export cover the surrounding pipeline. Search the [command reference](/api/powershell/) for `OfficeExcel`; use the [Excel examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/Excel) for end-to-end patterns.
