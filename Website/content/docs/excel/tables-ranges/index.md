---
title: Tables and Ranges
description: Excel tables with AutoFilter, data validation, and conditional formatting in OfficeIMO.Excel.
order: 22
---

# Tables and Ranges

OfficeIMO.Excel supports structured Excel tables with AutoFilter, built-in table styles, data validation rules, and conditional formatting. These features help create professional, interactive spreadsheets.

## Creating a Table

Tables in Excel provide structured references, automatic filtering, and styling. Create a table by defining the data range and applying a style:

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("tables.xlsx");
var sheet = workbook.AddWorkSheet("Sales");

// Populate data
sheet.Cells["A1"].Value = "Product";
sheet.Cells["B1"].Value = "Q1";
sheet.Cells["C1"].Value = "Q2";
sheet.Cells["A2"].Value = "Widget A";
sheet.Cells["B2"].Value = 15000;
sheet.Cells["C2"].Value = 18000;
sheet.Cells["A3"].Value = "Widget B";
sheet.Cells["B3"].Value = 22000;
sheet.Cells["C3"].Value = 25000;

// Create a table over the data range
sheet.AddTable("SalesTable", "A1:C3", TableStyle.Medium2);

workbook.Save();
```

## Table Styles

OfficeIMO provides access to all of Excel's built-in table styles through the `TableStyle` enum:

| Category | Examples |
|----------|---------|
| Light | `TableStyle.Light1` through `TableStyle.Light21` |
| Medium | `TableStyle.Medium1` through `TableStyle.Medium28` |
| Dark | `TableStyle.Dark1` through `TableStyle.Dark11` |

```csharp
sheet.AddTable("MyTable", "A1:D10", TableStyle.Dark3);
```

## Table Name Validation

Table names must be unique within a workbook. OfficeIMO validates this automatically:

```csharp
// The workbook tracks table names for uniqueness
workbook.TableNameComparer = StringComparer.OrdinalIgnoreCase;
```

You can control validation behavior:

```csharp
sheet.AddTable("MyTable", "A1:C5", TableStyle.Medium2,
    validationMode: TableNameValidationMode.ThrowOnInvalid);
```

## AutoFilter

Tables automatically include AutoFilter (drop-down filters on each column header). When you create a table, filtering is enabled by default.

For standalone AutoFilter without a formal table:

```csharp
sheet.SetAutoFilter("A1:C10");
```

## Fluent Table Builder

The fluent API provides a declarative way to define tables:

```csharp
using var workbook = ExcelFluentWorkbook.Create("fluent-table.xlsx")
    .Sheet("Data", s => s
        .Row(r => r.Cell("Name").Cell("Score").Cell("Grade"))
        .Row(r => r.Cell("Alice").Cell(95).Cell("A"))
        .Row(r => r.Cell("Bob").Cell(82).Cell("B"))
        .Row(r => r.Cell("Carol").Cell(91).Cell("A"))
        .Table("Grades", style: TableStyle.Medium9)
    )
    .Build();

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
