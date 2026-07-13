# OfficeIMO.Markup.Excel - Markup to Excel export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Markup.Excel)](https://www.nuget.org/packages/OfficeIMO.Markup.Excel)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markup.Excel?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Markup.Excel)

`OfficeIMO.Markup.Excel` exports the semantic `OfficeIMO.Markup` workbook model to editable Excel `.xlsx` files through `OfficeIMO.Excel`.

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

result.Document.SaveAsExcel("revenue.xlsx", new MarkupToExcelOptions {
});
```

## What it exports

- Sheets, sheet-qualified ranges, and formulas.
- Named tables and styled cell formatting.
- Dashboard charts from inline CSV data, worksheet ranges, or sheet-qualified table sources.
- Safe workbook defaults including gridline handling, table header freeze panes, auto-fit columns, defined-name repair, and Open XML validation controls.

## Boundaries

- Markup parsing and validation stay in `OfficeIMO.Markup`.
- Workbook creation and save behavior stay in `OfficeIMO.Excel`.
- This package maps semantic workbook nodes into editable Excel output.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Markup` and `OfficeIMO.Excel`; the exporter maps semantic nodes to editable workbook content.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
