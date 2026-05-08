# OfficeIMO.Markup.Excel

`OfficeIMO.Markup.Excel` exports the semantic `OfficeIMO.Markup` workbook model to editable Excel `.xlsx` files through `OfficeIMO.Excel`.

Use it when a `.omd` or Markdown-inspired authoring file has `profile: workbook` and should become a native workbook with sheets, ranges, formulas, tables, formatting, and charts.

## Install

```powershell
dotnet add package OfficeIMO.Markup.Excel
```

## Quick start

```csharp
using OfficeIMO.Markup;
using OfficeIMO.Markup.Excel;

var result = OfficeMarkupParser.Parse("""
---
profile: workbook
title: Revenue Workbook
---

@sheet {
  name: Revenue
}

::range address=A1
Product,2024,2025
A,100,120
B,80,92

::table name="RevenueTable" range=A1:C3 header=true

::formula cell=D2
=C2-B2
""");

new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
    OutputPath = "revenue.xlsx"
});
```

## What exports today

- Sheets, sheet-qualified ranges, and formulas
- Named tables and styled cell formatting
- Dashboard charts from inline CSV data, worksheet ranges, or sheet-qualified table sources
- Safe workbook defaults including gridline handling, table header freeze panes, auto-fit columns, defined-name repair, and Open XML validation controls

## Related packages

- `OfficeIMO.Markup`: parser, semantic AST, validation, and emitters
- `OfficeIMO.Markup.Cli`: command-line parse, validate, emit, and export workflow
- `OfficeIMO.Excel`: Excel workbook object model used by this exporter

## Targets

- `netstandard2.0`, `net8.0`, `net10.0`
- `net472` when building on Windows

License: MIT
