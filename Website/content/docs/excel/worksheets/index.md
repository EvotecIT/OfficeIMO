---
title: Worksheets
description: Working with Excel worksheets in OfficeIMO.Excel -- cells, formatting, formulas, and named ranges.
order: 21
---

# Worksheets

The `ExcelSheet` class represents a single worksheet within an Excel workbook. It provides access to cells, rows, columns, formatting, formulas, named ranges, and page setup.

## Creating Worksheets

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("worksheets.xlsx");

// Add worksheets
var sheet1 = workbook.AddWorkSheet("Sales");
var sheet2 = workbook.AddWorkSheet("Inventory");
var sheet3 = workbook.AddWorkSheet("Summary");

workbook.Save();
```

### Sheet Name Validation

By default, OfficeIMO validates sheet names against Excel's rules (no special characters like `:`, `\`, `/`, `?`, `*`, `[`, `]`; maximum 31 characters). You can control validation behavior:

```csharp
var sheet = workbook.AddWorkSheet("My Sheet!", SheetNameValidationMode.Sanitize);
// Name will be sanitized to remove invalid characters
```

## Cell Values

Access cells using A1-style references:

```csharp
var sheet = workbook.AddWorkSheet("Data");

// String values
sheet.Cells["A1"].Value = "Name";
sheet.Cells["B1"].Value = "Score";

// Numeric values
sheet.Cells["A2"].Value = "Alice";
sheet.Cells["B2"].Value = 95.5;

// Date values
sheet.Cells["A3"].Value = "Start Date";
sheet.Cells["B3"].Value = DateTime.Now;

// Boolean values
sheet.Cells["A4"].Value = "Active";
sheet.Cells["B4"].Value = true;
```

## Cell Formatting

### Number Formats

```csharp
// Currency format
sheet.Cells["B2"].NumberFormat = "$#,##0.00";

// Date format
sheet.Cells["B3"].NumberFormat = "yyyy-MM-dd";

// Percentage
sheet.Cells["B4"].NumberFormat = "0.00%";

// Use built-in presets
sheet.Cells["B5"].NumberPreset = ExcelNumberPreset.Currency;
```

### Font Styling

```csharp
sheet.Cells["A1"].Bold = true;
sheet.Cells["A1"].Italic = true;
sheet.Cells["A1"].FontSize = 14;
sheet.Cells["A1"].FontFamily = "Arial";
sheet.Cells["A1"].FontColor = "FF0000";  // Red
```

### Cell Fills and Borders

```csharp
sheet.Cells["A1"].BackgroundColor = "4472C4";  // Blue background
sheet.Cells["A1"].ForegroundColor = "FFFFFF";   // White text
```

### Alignment

```csharp
sheet.Cells["A1"].HorizontalAlignment = HorizontalAlignmentValues.Center;
sheet.Cells["A1"].VerticalAlignment = VerticalAlignmentValues.Center;
sheet.Cells["A1"].WrapText = true;
```

## Formulas

Set formulas using the standard Excel formula syntax:

```csharp
sheet.Cells["A1"].Value = 100;
sheet.Cells["A2"].Value = 200;
sheet.Cells["A3"].Value = 300;

// SUM formula
sheet.Cells["A4"].Value = "=SUM(A1:A3)";

// AVERAGE formula
sheet.Cells["A5"].Value = "=AVERAGE(A1:A3)";

// IF formula
sheet.Cells["B1"].Value = "=IF(A1>150,\"High\",\"Low\")";

// VLOOKUP, COUNTIF, and other functions work the same way
sheet.Cells["C1"].Value = "=COUNTIF(A1:A3,\">150\")";
```

## Column Width and Row Height

```csharp
// Set column width
sheet.SetColumnWidth("A", 25);
sheet.SetColumnWidth("B", 15);

// Auto-fit columns (requires data to be populated first)
sheet.AutoFitColumn("A");
```

## Named Ranges

Create and manage named ranges for use in formulas and references:

```csharp
// Define a named range
sheet.AddNamedRange("SalesData", "A1:C10");

// Use the named range in a formula
sheet.Cells["D1"].Value = "=SUM(SalesData)";
```

## Page Setup

Configure print settings:

```csharp
sheet.PageSetup.Orientation = OrientationValues.Landscape;
sheet.PageSetup.PaperSize = 1;  // Letter
sheet.PageSetup.FitToWidth = 1;
sheet.PageSetup.FitToHeight = 0;  // 0 = as many pages as needed
```

## Sheet Headers and Footers

```csharp
sheet.HeaderFooter.OddHeader = "&CMonthly Sales Report";
sheet.HeaderFooter.OddFooter = "&LPage &P of &N&R&D";
```

## Images

Add images to worksheets:

```csharp
sheet.AddImage("chart.png", "A10", width: 400, height: 300);
```

## Freezing Panes

```csharp
// Freeze the top row
sheet.FreezePane(1, 0);

// Freeze the first column
sheet.FreezePane(0, 1);

// Freeze both (top-left corner stays fixed)
sheet.FreezePane(1, 1);
```

## Iterating Over Sheets

```csharp
using var workbook = ExcelDocument.Load("data.xlsx");

foreach (var sheet in workbook.Sheets) {
    Console.WriteLine($"Sheet '{sheet.Name}' has data in {sheet.RowCount} rows");
}
```
