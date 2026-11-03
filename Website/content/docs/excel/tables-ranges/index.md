---
title: Tables and Ranges
description: Excel tables with AutoFilter, data validation, and conditional formatting in OfficeIMO.Excel. Includes examples and API links.
order: 22
---

OfficeIMO.Excel supports structured Excel tables with AutoFilter, built-in table styles, data validation rules, and conditional formatting. These features help create professional, interactive spreadsheets.

## Creating a Table

Tables in Excel provide structured references, automatic filtering, and styling. Create a table by defining the data range and applying a style:

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("tables.xlsx");
var sheet = workbook.AddSheet("Sales");

// Populate data
sheet.Cell("A1").SetValue("Product");
sheet.Cell("B1").SetValue("Q1");
sheet.Cell("C1").SetValue("Q2");
sheet.Cell("A2").SetValue("Widget A");
sheet.Cell("B2").SetValue(15000);
sheet.Cell("C2").SetValue(18000);
sheet.Cell("A3").SetValue("Widget B");
sheet.Cell("B3").SetValue(22000);
sheet.Cell("C3").SetValue(25000);

// Create a table over the data range
sheet.AddTable("A1:C3", name: "SalesTable", style: ExcelTableStyle.TableStyleMedium2);

workbook.Save();
```

## Table Styles

OfficeIMO provides access to all of Excel's built-in table styles through the `ExcelTableStyle` enum:

| Category | Examples |
|----------|---------|
| Light | `ExcelTableStyle.TableStyleLight1` through `ExcelTableStyle.TableStyleLight21` |
| Medium | `ExcelTableStyle.TableStyleMedium1` through `ExcelTableStyle.TableStyleMedium28` |
| Dark | `ExcelTableStyle.TableStyleDark1` through `ExcelTableStyle.TableStyleDark11` |

```csharp
sheet.AddTable("MyTable", "A1:D10", ExcelTableStyle.TableStyleDark3);
```

## Table Name Validation

Table names must be unique within a workbook. OfficeIMO validates this automatically:

```csharp
// The workbook tracks table names for uniqueness
workbook.TableNameComparer = StringComparer.OrdinalIgnoreCase;
```

You can control validation behavior:

```csharp
sheet.AddTable("MyTable", "A1:C5", ExcelTableStyle.TableStyleMedium2,
    validationMode: TableNameValidationMode.ThrowOnInvalid);
```

## AutoFilter

Tables automatically include AutoFilter (drop-down filters on each column header). When you create a table, filtering is enabled by default.

For standalone AutoFilter without a formal table, add it to the range and set criteria by header:

```csharp
sheet.Range("A1:C10").AutoFilter()
    .Column("Region").Equals("North", "South")
    .Column("Amount").GreaterThanOrEqual(1000);
```

## Add a typed table

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("grades.xlsx");
var sheet = workbook.AddSheet("Data");
var table = sheet.AddTable(new[] {
    new { Name = "Alice", Score = 95, Grade = "A" },
    new { Name = "Bob", Score = 82, Grade = "B" },
    new { Name = "Carol", Score = 91, Grade = "A" }
}, new ExcelTableOptions { Name = "Grades", Style = ExcelTableStyle.TableStyleMedium9 });
workbook.Save();
```

## Data Validation

Apply data validation rules to cells to restrict user input:

```csharp
// Dropdown list validation
sheet.Range("B2:B100").Validation().List("High", "Medium", "Low");

// Numeric range validation
sheet.Range("C2:C100").Validation().WholeNumber(ExcelDataValidationOperator.Between, 0, 100,
    errorMessage: "Value must be between 0 and 100");

// Date validation
sheet.Range("D2:D100").Validation().Date(ExcelDataValidationOperator.Between,
    new DateTime(2025, 1, 1), new DateTime(2025, 12, 31));
```

## Conditional Formatting

Apply visual formatting rules based on cell values:

```csharp
// Highlight cells greater than a threshold
sheet.Range("B2:B100").ConditionalFormat()
    .CellIs(ExcelConditionalFormattingOperator.GreaterThan, "10000", fillColor: "92D050");

// Highlight cells with specific text
sheet.Range("A2:A100").ConditionalFormat()
    .Text(ExcelConditionalFormatType.ContainsText, "Urgent", fillColor: "FF0000");

// Color scale (gradient from red to green)
sheet.Range("C2:C100").ConditionalFormat().ColorScale("FF0000", "00FF00");

// Data bars
sheet.Range("D2:D100").ConditionalFormat().DataBar("4472C4");

// Icon sets
sheet.Range("E2:E100").ConditionalFormat().IconSet(ExcelIconSet.ThreeArrows);
```

Each call returns the rule it added. List, update, reorder and remove existing rules with
`sheet.GetConditionalFormattingRules`, `UpdateConditionalFormattingRule`, `ReorderConditionalFormattingRules`,
`RemoveConditionalFormattingRule` and `ClearConditionalFormatting`.

## DataTable Import

Import data from a `DataTable` directly into a sheet:

```csharp
var dataTable = new System.Data.DataTable();
dataTable.Columns.Add("Name", typeof(string));
dataTable.Columns.Add("Value", typeof(double));
dataTable.Rows.Add("Item A", 100.5);
dataTable.Rows.Add("Item B", 200.75);

sheet.ImportDataTable(dataTable, startCell: "A1", includeHeaders: true);
```

## Reading Tables

When loading an existing workbook, access defined tables:

```csharp
using var workbook = ExcelDocument.Load("existing.xlsx");
var sheet = workbook.Sheets[0];

foreach (var table in sheet.Tables) {
    Console.WriteLine($"Table: {table.Name}, Range: {table.Reference}");
}
```
