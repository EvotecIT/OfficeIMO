---
title: "Read and Import Excel Data"
description: "Read worksheets, ranges, tables, and typed rows from existing workbooks without requiring Microsoft Excel."
layout: docs
---

Use `Import-OfficeExcel` when the workbook is a data source. Use `Open-OfficeExcel` and the targeted inspection commands when workbook structure, formatting, formulas, validation, charts, or metadata matter.

## Import rows

```powershell
$rows = Import-OfficeExcel -Path '.\Input\Register.xlsx' `
    -Sheet 'Services' -Range 'A1:D500'
$atRisk = $rows | Where-Object Status -eq 'At risk'
```

A wildcard `-Sheet` reads several worksheets and adds sheet identity to each result. `-As Hashtable`, `-As DataTable`, `-As DataReader` and `-As Column` return other shapes for different downstream consumers. `-As DataReader` reads one worksheet at a time, so pick a single `-Sheet`, and the caller disposes the reader it returns. Range and row/column bounds keep large imports intentional.

## Inspect workbook structure

Use `Get-OfficeDocumentInfo -Detailed` for a workbook-level inventory, then query tables, named ranges, formulas, comments, validation, pivots, worksheet views, page breaks, links, queries, and rich text only when needed.

The [read-and-filter recipe](https://github.com/EvotecIT/OfficeIMO/blob/0b3d4a4b27636e4908697be77b80a15b4b5e3a7b/PowerShell/PSWriteOffice/Examples/Excel/Recipe-Excel-ReadAndFilter.ps1) creates a table, imports typed rows, and filters the records that require attention.

For semicolon-delimited or other text data, continue with [delimited import and publishing](/docs/pswriteoffice/excel-export-publish/).
