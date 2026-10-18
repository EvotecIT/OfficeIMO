---
title: "Import, Export, and Publish Excel Data"
description: "Move delimited data into Excel, export workbook visuals and HTML, and choose the right publishing surface for downstream readers."
layout: docs
---

Excel often sits in the middle of a pipeline: CSV or application data arrives, a workbook adds formulas and presentation, and selected results are published as images, HTML, or another workbook.

## Import delimited text

`Import-OfficeExcelDelimitedText` adds normalized CSV or other delimited data to an existing workbook. Specify the delimiter and culture rather than relying on machine defaults.

```powershell
Import-OfficeExcelDelimitedText -Path '.\Report.xlsx' `
    -SourcePath '.\Sales.csv' -Delimiter ';' `
    -CultureName 'en-US' -SheetName Sales
```

The [delimited-import recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-ImportDelimited.ps1) imports an existing semicolon-delimited file into a named worksheet.

## Publish the useful part

Use `Export-OfficeExcelRangeImage` for a bounded report region, `Export-OfficeExcelChartImage` for a chart, and `Export-OfficeExcelImage` for workbook-oriented image output. HTML export is useful for review without Excel. `Export-OfficeExcel` turns PowerShell objects into a new or appended workbook when the pipeline begins with objects rather than an existing file.

Choose the smallest artifact that preserves the intended experience. A range image is good for a status message; an XLSX is better when recipients must filter, inspect formulas, or continue editing.
