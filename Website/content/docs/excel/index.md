---
title: Excel Workbooks
description: Overview of the OfficeIMO.Excel package for creating and manipulating Excel spreadsheets. Includes examples and API links.
order: 20
---

The `OfficeIMO.Excel` package provides a higher-level API for creating, reading, and modifying Excel workbooks (`.xlsx`) without requiring Microsoft Office. It uses the Open XML SDK, an editable object model, a report builder, and a streaming writer for large exports.

## Key Classes

| Class | Description |
|-------|-------------|
| `ExcelDocument` | Root class for creating and loading workbooks. Implements `IDisposable` and `IAsyncDisposable`. |
| `ExcelSheet` | Represents a single worksheet with cells, rows, columns, tables, and conditional formatting. |
| `ExcelTable` | An editable structured table with filters and styles. |
| `ExcelReportBuilder` | Sections, KPI rows, tables, and other report elements, reached through `sheet.Report()`. |
| `ExcelWorkbookWriter` | A bounded writer for forward-only exports, created by `ExcelDocument.CreateWriter()`. |

## Creating a Workbook

```csharp
using OfficeIMO.Excel;

// Create with a file path
using var workbook = ExcelDocument.Create("report.xlsx");
var sheet = workbook.AddWorksheet("Sales");

sheet.CellAt("A1").SetValue("Product");
sheet.CellAt("B1").SetValue("Revenue");
sheet.CellAt("A2").SetValue("Widget A");
sheet.CellAt("B2").SetValue(15000);

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
var sheet = workbook.AddWorksheet("Data");
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

## Add typed data

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("monthly.xlsx");
workbook.BuiltinDocumentProperties.Title = "Monthly Report";
workbook.BuiltinDocumentProperties.Creator = "OfficeIMO";
var summary = workbook.AddWorksheet("Summary");
summary.AddTable(new[] {
    new { Metric = "Users", Value = 1250 },
    new { Metric = "Revenue", Value = 45000 }
}, new ExcelTableOptions { Name = "Metrics" });
var details = workbook.AddWorksheet("Details");
details.AddTable(new[] { new { Item = "Widget A", Quantity = 100, Price = 15.99m } });
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
workbook.Execution.Mode = ExcelExecutionMode.Parallel;
workbook.Execution.MaxDegreeOfParallelism = 4;
```

## Sheet Caching

For very large workbooks, you can disable sheet wrapper caching to reduce memory:

```csharp
workbook.SheetCachingEnabled = false;
```

## CSV interoperability

Install the optional `OfficeIMO.Excel.Csv` adapter when a workflow needs to move data between CSV and Excel. The adapter keeps parsing and writing in `OfficeIMO.CSV`, workbook operations in `OfficeIMO.Excel`, and avoids adding CSV dependencies to Excel-only or Reader workloads.

```powershell
dotnet add package OfficeIMO.Excel.Csv
```

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;

using var workbook = ExcelDocument.Create();
ExcelCsvImportResult imported = workbook.ImportCsvFile("sales.csv");
workbook[imported.SheetName].SaveAsCsv("sales-roundtrip.csv");
workbook.Save("sales.xlsx");
```

`imported.Delimiter` reports the delimiter that was actually used, including
one selected automatically by delimiter detection. `imported.TableName` reports
the sanitized, unique table name that was actually created.

Use `OfficeIMO.Reader.Excel` instead when the job is normalized extraction of workbook content for search, indexing, or AI ingestion; it does not own CSV conversion.

## Further Reading

- [Worksheets](/docs/excel/worksheets) — Cell values, formatting, formulas, and named ranges.
- [Tables and Ranges](/docs/excel/tables-ranges) — AutoFilter tables, validation, and conditional formatting.
