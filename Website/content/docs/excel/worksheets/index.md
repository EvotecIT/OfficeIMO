---
title: Worksheets
description: Working with Excel worksheets in OfficeIMO.Excel — cells, formatting, formulas, and named ranges. Includes examples and API links.
order: 21
---

The `ExcelSheet` class represents a single worksheet within an Excel workbook. It provides access to cells, rows, columns, formatting, formulas, named ranges, and page setup.

## Creating Worksheets

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("worksheets.xlsx");

// Add worksheets
var sheet1 = workbook.AddSheet("Sales");
var sheet2 = workbook.AddSheet("Inventory");
var sheet3 = workbook.AddSheet("Summary");

workbook.Save();
```

### Sheet Name Validation

By default, OfficeIMO validates sheet names against Excel's rules (no special characters like `:`, `\`, `/`, `?`, `*`, `[`, `]`; maximum 31 characters). You can control validation behavior:

```csharp
var sheet = workbook.AddSheet("My Sheet!", ExcelSheetNameValidationMode.Sanitize);
// Name will be sanitized to remove invalid characters
```

## Cell Values

Access cells using A1-style references:

```csharp
var sheet = workbook.AddSheet("Data");

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
sheet.Cells["A1"].HorizontalAlignment = ExcelHorizontalAlignment.Center;
sheet.Cells["A1"].VerticalAlignment = ExcelVerticalAlignment.Center;
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
sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(25));
sheet.SetColumnWidth(2, ExcelColumnWidth.Length(OfficeLength.Centimeters(3)));
sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeLength.Points(24)));

// Auto-fit columns (requires data to be populated first)
sheet.AutoFitColumn(1);
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
sheet.PageSetup.Orientation = OfficePageOrientation.Landscape;
sheet.PageSetup.PaperSize = ExcelPaperSize.Letter;
sheet.PageSetup.FitToWidth = 1;
sheet.PageSetup.FitToHeight = 0;  // 0 = as many pages as needed
sheet.PageSetup.SetMargins(ExcelMarginPreset.Narrow);
sheet.PageSetup.PrintArea = "A1:F40";
sheet.PageSetup.SetPrintTitles(firstRow: 1, lastRow: 1);
```

## Sheet Headers and Footers

Each header and footer has left, center and right sections. Excel codes such as `&P` (page), `&N` (pages) and `&D` (date) are kept.

```csharp
sheet.PageSetup.Header.Center.Text = "Monthly Sales Report";
sheet.PageSetup.Footer.Left.Text = "Page &P of &N";
sheet.PageSetup.Footer.Right.Text = "&D";
sheet.PageSetup.Header.Right.SetImage(File.ReadAllBytes("logo.png"), "image/png");

// A different first page
sheet.PageSetup.FirstPageHeader.Center.Text = "Monthly Sales Report - cover";
```

## Images

Add images to worksheets:

```csharp
sheet.AddImageFromFile(10, 1, "chart.png", width: OfficeIMO.OfficeLength.Pixels(400), height: OfficeIMO.OfficeLength.Pixels(300));
```

## Freezing Panes

```csharp
// Freeze the top row
sheet.View.Freeze(topRows: 1);

// Freeze the first column
sheet.View.Freeze(leftColumns: 1);

// Freeze both (top-left corner stays fixed)
sheet.View.Freeze(1, 1);
```

Gridlines, zero values, row and column headings, direction and zoom are on `sheet.View` too:

```csharp
sheet.View.ShowGridlines = false;
sheet.View.ZoomScale = 90;
```

## Iterating Over Sheets

```csharp
using var workbook = ExcelDocument.Load("data.xlsx");

foreach (var sheet in workbook.Sheets) {
    Console.WriteLine($"Sheet '{sheet.Name}' has data in {sheet.RowCount} rows");
}
```
