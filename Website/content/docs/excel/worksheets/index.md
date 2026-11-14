---
title: Worksheets
description: Working with Excel worksheets in OfficeIMO.Excel — cells, formatting, formulas, and named ranges. Includes examples and API links.
order: 21
search_terms:
  - add sheet
  - excel worksheet
  - freeze panes
  - column width
---

The `ExcelSheet` class represents a single worksheet within an Excel workbook. It provides access to cells, ranges, rows, columns, formatting, formulas, named ranges, page setup, and the sheet view.

Rows and columns are 1-based, as in Excel. `sheet.Cell("B4")` and `sheet.Cell(4, 2)` return a lightweight `ExcelCell` handle, and `sheet.Range("A1:D1")` returns an `ExcelRange`. Cell and range methods return the same handle, so calls can be chained. The examples after the first one assume a `sheet` created as shown there.

## Creating Worksheets

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("worksheets.xlsx");

// Add worksheets
var sales = workbook.AddSheet("Sales");
var inventory = workbook.AddSheet("Inventory");
var summary = workbook.AddSheet("Summary");

workbook.Save();
```

### Sheet Name Validation

By default, OfficeIMO validates sheet names against Excel's rules (no special characters like `:`, `\`, `/`, `?`, `*`, `[`, `]`; maximum 31 characters). You can control validation behavior:

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("names.xlsx");
var sheet = workbook.AddSheet("Q1/Q2 [draft]", ExcelSheetNameValidationMode.Sanitize);
// Invalid characters are removed from the name
Console.WriteLine(sheet.Name);
```

## Cell Values

Address cells with A1 references or with row and column numbers. `SetValue` keeps the .NET type: numbers stay numbers, dates become Excel dates, and booleans become `TRUE`/`FALSE`.

```csharp
// Text
sheet.Cell("A1").SetValue("Name");
sheet.Cell("B1").SetValue("Score");

// Numbers
sheet.Cell("A2").SetValue("Alice");
sheet.Cell("B2").SetValue(95.5);

// Dates
sheet.Cell("A3").SetValue("Start Date");
sheet.Cell(3, 2).SetValue(new DateTime(2026, 11, 2));

// Booleans
sheet.Cell("A4").SetValue("Active");
sheet.Cell("B4").SetValue(true);

// Several values at once
sheet.Range("D1:E2").SetValues(new object?[,] {
    { "Region", "Revenue" },
    { "North", 27_500m }
});

// Read values back
double score = sheet.Cell("B2").GetValue<double>();
string text = sheet.Cell("B3").GetFormattedText();
```

## Cell Formatting

Format a block of cells through its range rather than cell by cell: one range call applies the style once.

### Number Formats

Number formats use Excel's format codes:

```csharp
// Currency
sheet.Cell("B2").SetNumberFormat("$#,##0.00");

// Date
sheet.Cell("B3").SetNumberFormat("yyyy-mm-dd");

// Percentage, applied to a whole range
sheet.Range("C2:C10").SetNumberFormat("0.00%");
```

For columns written from objects, `sheet.ColumnStyleByHeader("Revenue").Currency(decimals: 2)` applies a preset by header name. See [Styles and Number Formats](/docs/excel/styles-number-formats/).

### Font Styling

```csharp
using OfficeIMO;

sheet.Cell("A1")
    .SetBold()
    .SetItalic()
    .SetFontSize(OfficeLength.Points(14))
    .SetFontFamily("Arial")
    .SetColor(OfficeColor.Parse("FF0000"));  // Red text
```

### Cell Fills and Borders

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

sheet.Range("A1:B1")
    .SetFillColor(OfficeColor.Parse("4472C4"))   // Blue background
    .SetColor(OfficeColor.Parse("FFFFFF"))       // White text
    .SetBorder(ExcelBorderStyle.Thin, OfficeColor.Parse("1F3864"));
```

### Alignment

```csharp
using OfficeIMO.Excel;

sheet.Cell("A1")
    .SetHorizontalAlignment(ExcelHorizontalAlignment.Center)
    .SetVerticalAlignment(ExcelVerticalAlignment.Center)
    .SetWrapText();
```

## Formulas

Set formulas with `SetFormula`, using standard Excel syntax without the leading `=`:

```csharp
sheet.Cell("A1").SetValue(100);
sheet.Cell("A2").SetValue(200);
sheet.Cell("A3").SetValue(300);

// SUM formula
sheet.Cell("A4").SetFormula("SUM(A1:A3)");

// AVERAGE formula
sheet.Cell("A5").SetFormula("AVERAGE(A1:A3)");

// IF formula
sheet.Cell("B1").SetFormula("IF(A1>150,\"High\",\"Low\")");

// VLOOKUP, COUNTIF, and other functions work the same way
sheet.Cell("C1").SetFormula("COUNTIF(A1:A3,\">150\")");
```

Call `workbook.Calculate()` to compute the supported functions and cache their results. See [Formulas and Named Ranges](/docs/excel/formulas/).

## Column Width and Row Height

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

// Set column width
sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(25));
sheet.SetColumnWidth(2, ExcelColumnWidth.Length(OfficeLength.Centimeters(3)));
sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeLength.Points(24)));

// Auto-fit (requires data to be populated first)
sheet.AutoFitColumn(1);
sheet.AutoFitColumns();
```

## Named Ranges

Create named ranges for use in formulas and references. A sheet-level name is visible only on that sheet; a workbook-level name is visible everywhere.

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("named.xlsx");
var sheet = workbook.AddSheet("Data");
sheet.Range("A1:C10").SetNumberFormat("#,##0");

// Workbook-level name
workbook.SetNamedRange("SalesData", "'Data'!A1:C10");

// Sheet-level name
sheet.SetNamedRange("FirstValue", "A1");

// Use the named range in a formula
sheet.Cell("D1").SetFormula("SUM(SalesData)");

workbook.Save();
```

## Page Setup

Configure print settings:

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

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
sheet.PageSetup.DifferentFirstPage = true;
sheet.PageSetup.FirstPageHeader.Center.Text = "Monthly Sales Report - cover";
```

## Images

Add an image anchored at a cell:

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

sheet.AddImage("A10", "chart.png", new ExcelImageOptions {
    Width = OfficeLength.Pixels(400),
    Height = OfficeLength.Pixels(300)
});
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
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Load("data.xlsx");

foreach (var sheet in workbook.Sheets) {
    Console.WriteLine($"Sheet '{sheet.Name}' uses {sheet.UsedRange.Address}");
}
```
