---
title: "Automate Excel Workbooks"
description: "Build, inspect, validate, compare, repair, and publish workbook workflows from PowerShell. Includes examples and cmdlet links."
layout: docs
---

Excel is the largest PSWriteOffice family with 159 exported commands. It covers workbook creation and reading, sheet and range operations, formulas, styling, tables, charts, pivots, validation, comments, images, links, templates, dashboards, protection, accessibility, comparison, repair, streaming contracts, and direct range or chart image export.

## Export rows in one line

Use the pipeline when the job is simply PowerShell objects to a worksheet:

```powershell
$records | Export-OfficeExcel -Path '.\Revenue.xlsx' -WorksheetName 'Sales' -TableName 'Sales'
```

Use the object or DSL surface when you need several sheets, formulas, charts, pivots, or precise formatting.

## Create a workbook from data

Use `New-OfficeExcel` with `Add-OfficeExcelSheet`, then add tables, formulas, charts, and report components inside each sheet context. The report DSL includes titles, paragraphs, sections, callouts, KPI rows, tables, legends, spacers, and dashboard charts for repeatable operational output.

```powershell
$records = @(
    [pscustomobject]@{ Region = 'EMEA'; Revenue = 98000 }
    [pscustomobject]@{ Region = 'APAC'; Revenue = 143000 }
)

New-OfficeExcel -Path '.\Output\Revenue.xlsx' {
    Add-OfficeExcelSheet -Name 'Sales' {
        Add-OfficeExcelTable -InputObject $records -TableName 'Sales' -AutoFit
        Add-OfficeExcelChart -TableName 'Sales' -Row 2 -Column 5 -Type ColumnClustered -Title 'Revenue by region'
    }
}
```

## Copy complete DSL recipes

- [Project tracker](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-ProjectTracker.ps1): table, validation list, conditional rules, chart, frozen header, print layout, and workbook index.
- [Budget dashboard](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-BudgetDashboard.ps1): summary formulas and chart over a separate styled detail sheet.
- [Operational dashboard](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1): the larger example with KPIs, pivots, sparklines, charts, links, comments, queries, and validation.

## Other Excel workflows

- [Object composition](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-ObjectComposition.ps1): keep a workbook and worksheet object, then add cells and a table incrementally.
- [Append or replace rows](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-AppendAndReplace.ps1): update a recurring export without rebuilding the surrounding workbook.
- [Template invoice](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-TemplateInvoice.ps1): fill a prepared workbook while preserving its design.
- [Pivots and sparklines](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-PivotAndSparklines.ps1): build analytical views beyond a flat export.

The [DSL cookbook](/docs/pswriteoffice/dsl-cookbook/) shows how Excel composition differs from the Word, PowerPoint, PDF, and Markdown blocks that consume the same objects.

## Task guides

- [Read and import data](/docs/pswriteoffice/excel-read-import/)
- [Update existing workbooks](/docs/pswriteoffice/excel-update-existing/)
- [Merge and compare workbooks](/docs/pswriteoffice/excel-merge-compare/)
- [Validate and repair workbooks](/docs/pswriteoffice/excel-validation-repair/)
- [Import, export, and publish](/docs/pswriteoffice/excel-export-publish/)

## Work with existing workbooks

The read surface can return used ranges, tables, named ranges, formulas, comments, validation, rich text, worksheet views, page breaks, pivots, summaries, preflight data, and streaming capabilities. Targeted commands update cells, rows, columns, styles, formulas, links, page setup, print settings, themes, worksheet visibility, active sheet, filters, and write reservations.

## Validate before delivery

- `Get-OfficeExcelPreflight` and `Get-OfficeExcelRuntimePreflight` report readiness before an operation.
- `Test-OfficeExcelWorkbook` checks workbook integrity.
- `Test-OfficeExcelAccessibility` supports accessible-delivery gates.
- `Compare-OfficeExcelWorkbook` and `Compare-OfficeExcelRange` make change evidence explicit.
- `Repair-OfficeExcelWorkbook` is a deliberate repair path, not an implicit side effect of reading.

Templates, joins, merges, sheet ordering, workbook copying, HTML review, and delimited import/export cover the surrounding pipeline. Search the [command reference](/api/powershell/) for `OfficeExcel`; use the [Excel examples](https://github.com/EvotecIT/PSWriteOffice/tree/main/Examples/Excel) for end-to-end patterns.
