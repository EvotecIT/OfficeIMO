---
title: "Convert XLS and XLSX in .NET"
description: "Modernize Excel 97–2003 XLS workbooks or create supported XLS output from XLSX with OfficeIMO.Excel compatibility diagnostics."
meta.eyebrow: "Excel conversion"
meta.source_format: "XLS"
meta.destination_format: "XLSX"
meta.package: "OfficeIMO.Excel"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Excel"
meta.runtime: ".NET on Windows, Linux, and macOS"
meta.howto.name: "Convert an XLS workbook to XLSX with OfficeIMO.Excel"
meta.howto.description: "Inspect an XLS workbook, enforce a compatibility policy, and write XLSX without Microsoft Excel."
meta.howto.steps:
  - "Install|Add the OfficeIMO.Excel NuGet package."
  - "Analyze|Preview formulas, styles, charts, and other compatibility findings."
  - "Convert|Call ExcelDocument.Convert with XLS and XLSX paths."
  - "Validate|Open or inspect the output required by the target workflow."
---

`OfficeIMO.Excel` reads legacy BIFF workbooks and modern Open XML workbooks through one Excel API. That makes it useful for finance archives, line-of-business exports, scheduled imports, and migrations where `.xls` still appears alongside `.xlsx`.

## XLS to XLSX

```csharp
using OfficeIMO.Excel;

ExcelDocumentConversionReport preview =
    ExcelDocument.AnalyzeConversion("forecast.xls", "forecast.xlsx");

ExcelDocumentConversionResult result =
    ExcelDocument.Convert("forecast.xls", "forecast.xlsx");
```

The conversion report describes how worksheets, formulas, formatting, charts, drawing objects, validation, names, macros, and other tracked capabilities are handled. A batch process can accept known approximations while blocking findings that would change business meaning.

## XLSX to XLS

```csharp
using OfficeIMO.Excel;

ExcelDocumentConversionResult result =
    ExcelDocument.Convert("forecast.xlsx", "forecast.xls");
```

The older XLS format has stricter limits and different feature models. Analyze before writing if the workbook uses modern tables, chart features, conditional formatting, rich drawings, or dimensions beyond legacy limits.

## A safe migration pattern

Keep the original file until the converted workbook passes the checks your workflow values: formula preservation, named ranges, worksheet dimensions, expected values, or a controlled desktop-Excel validation step. OfficeIMO can run without Excel; an optional Excel check can still act as an independent oracle for high-value migrations.

See [Excel compatibility](/compatibility/#excel) for the tracked support boundary and the [Excel guide](/docs/excel/) for editing after conversion.
