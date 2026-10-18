---
title: "Read and Import Excel Data"
description: "Read worksheets, ranges, tables, and typed rows from existing workbooks without requiring Microsoft Excel."
layout: docs
---

Use `Import-OfficeExcel` when the workbook is a data source. Use `Get-OfficeExcel` and the targeted inspection commands when workbook structure, formatting, formulas, validation, charts, or metadata matter.

## Import rows

```powershell
$rows = Import-OfficeExcel -Path '.\Input\Register.xlsx' `
    -WorksheetName 'Services' -Range 'A1:D500'
$atRisk = $rows | Where-Object Status -eq 'At risk'
```

`-AllSheets` adds sheet identity to each result. `-ByColumn`, `-AsHashtable`, `-AsDataTable`, and `-AsDataReader` support different downstream consumers. Range and row/column bounds keep large imports intentional.

## Inspect workbook structure

Use `Get-OfficeExcelSummary` for a workbook-level inventory, then query tables, named ranges, formulas, comments, validation, pivots, worksheet views, page breaks, links, queries, and rich text only when needed.

The [read-and-filter recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-ReadAndFilter.ps1) creates a table, imports typed rows, and filters the records that require attention.

For semicolon-delimited or other text data, continue with [delimited import and publishing](/docs/pswriteoffice/excel-export-publish/).
