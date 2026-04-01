---
title: Excel Workbooks
description: Overview of the OfficeIMO.Excel package for creating and manipulating Excel spreadsheets.
order: 20
---

# Excel Workbooks

The `OfficeIMO.Excel` package provides a higher-level API for creating, reading, and modifying Excel workbooks (`.xlsx`) without requiring Microsoft Office. It is built on the Open XML SDK and exposes both imperative and fluent patterns, with additional examples and tests in the repo covering larger reporting scenarios.

## Key Classes

| Class | Description |
|-------|-------------|
| `ExcelDocument` | Root class for creating and loading workbooks. Implements `IDisposable` and `IAsyncDisposable`. |
| `ExcelSheet` | Represents a single worksheet with cells, rows, columns, tables, and conditional formatting. |
| `ExcelFluentWorkbook` | Fluent builder for composing workbooks with chained method calls. |
| `SheetBuilder` | Fluent builder for configuring individual sheets. |
| `TableBuilder` | Fluent builder for defining Excel tables within a sheet. |
| `StyleBuilder` | Fluent builder for cell and range styling. |

## Creating a Workbook

```csharp
using OfficeIMO.Excel;

// Create with a file path
using var workbook = ExcelDocument.Create("report.xlsx");
var sheet = workbook.AddWorkSheet("Sales");

sheet.Cells["A1"].Value = "Product";
sheet.Cells["B1"].Value = "Revenue";
sheet.Cells["A2"].Value = "Widget A";
sheet.Cells["B2"].Value = 15000;

workbook.Save();
```

Create with a named worksheet in one call:

```csharp
using var workbook = ExcelDocument.Create("report.xlsx", "Sales");
var sheet = workbook.Sheets[0]; // "Sales" sheet
```

## Creating a Workbook in Memory

```csharp
using var stream = new MemoryStream();
using var workbook = ExcelDocument.Create(stream);
var sheet = workbook.AddWorkSheet("Data");
// ... populate ...
workbook.Save();
```

## Loading an Existing Workbook

```csharp
using var workbook = ExcelDocument.Load("existing.xlsx");

foreach (var sheet in workbook.Sheets) {
    Console.WriteLine($"Sheet: {sheet.Name}");
}
```

## Fluent API

The fluent API provides a declarative way to compose workbooks:

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelFluentWorkbook.Create("fluent.xlsx")
    .Info(i => i.Title("Monthly Report").Author("OfficeIMO"))
    .Sheet("Summary", s => s
        .Row(r => r.Cell("Metric").Cell("Value"))
        .Row(r => r.Cell("Users").Cell(1250))
        .Row(r => r.Cell("Revenue").Cell(45000))
    )
    .Sheet("Details", s => s
        .Row(r => r.Cell("Item").Cell("Qty").Cell("Price"))
        .Row(r => r.Cell("Widget A").Cell(100).Cell(15.99))
    )
    .Build();

workbook.Save();
```

## Document Properties

```csharp
using var workbook = ExcelDocument.Create("props.xlsx");

workbook.BuiltinDocumentProperties.Title = "Financial Report";
workbook.BuiltinDocumentProperties.Creator = "Finance Team";
workbook.ApplicationProperties.Company = "Evotec";

workbook.Save();
```

## Execution Policy

Control parallel vs sequential operations for large workbooks:

```csharp
var workbook = ExcelDocument.Create("large.xlsx");
workbook.Execution.Mode = ExecutionMode.Parallel;
workbook.Execution.MaxDegreeOfParallelism = 4;
```

## Sheet Caching

For very large workbooks, you can disable sheet wrapper caching to reduce memory:

```csharp
workbook.SheetCachingEnabled = false;
```

## Further Reading

- [Worksheets](/docs/excel/worksheets) -- Cell values, formatting, formulas, and named ranges.
- [Tables and Ranges](/docs/excel/tables-ranges) -- AutoFilter tables, validation, and conditional formatting.
