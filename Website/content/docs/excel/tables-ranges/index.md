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
var sheet = workbook.AddWorksheet("Sales");

// Populate data
sheet.CellAt("A1").SetValue("Product");
sheet.CellAt("B1").SetValue("Q1");
sheet.CellAt("C1").SetValue("Q2");
sheet.CellAt("A2").SetValue("Widget A");
sheet.CellAt("B2").SetValue(15000);
sheet.CellAt("C2").SetValue(18000);
sheet.CellAt("A3").SetValue("Widget B");
sheet.CellAt("B3").SetValue(22000);
sheet.CellAt("C3").SetValue(25000);

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

For standalone AutoFilter without a formal table:

```csharp
sheet.SetAutoFilter("A1:C10");
```

## Add a typed table

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("grades.xlsx");
var sheet = workbook.AddWorksheet("Data");
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
sheet.AddDataValidation("B2:B100",
    DataValidationType.List,
    formula: "\"High,Medium,Low\"");

// Numeric range validation
sheet.AddDataValidation("C2:C100",
    DataValidationType.Whole,
    minimum: "0",
    maximum: "100",
    errorMessage: "Value must be between 0 and 100");

// Date validation
sheet.AddDataValidation("D2:D100",
    DataValidationType.Date,
    minimum: "2025-01-01",
    maximum: "2025-12-31");
```

## Conditional Formatting

Apply visual formatting rules based on cell values:

```csharp
// Highlight cells greater than a threshold
sheet.AddConditionalFormatting("B2:B100",
    ConditionalFormattingRuleType.CellIs,
    operatorValue: ConditionalFormattingOperator.GreaterThan,
    formula: "10000",
    backgroundColor: "92D050");  // Green

// Highlight cells with specific text
sheet.AddConditionalFormatting("A2:A100",
    ConditionalFormattingRuleType.ContainsText,
    text: "Urgent",
    backgroundColor: "FF0000",   // Red background
    fontColor: "FFFFFF");        // White text

// Color scale (gradient from red to green)
sheet.AddColorScale("C2:C100",
    minColor: "FF0000",   // Red for low values
    maxColor: "00FF00");  // Green for high values

// Data bars
sheet.AddDataBars("D2:D100", color: "4472C4");

// Icon sets
sheet.AddIconSet("E2:E100", IconSetType.ThreeArrows);
```

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
