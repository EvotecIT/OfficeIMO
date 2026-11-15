---
title: Styles and Number Formats
description: Format Excel cells, ranges, and table columns with fonts, fills, borders, alignment, number formats, and conditional formatting in OfficeIMO.Excel.
order: 24
packages:
  - OfficeIMO.Excel
search_terms:
  - excel formatting
  - number format
  - cell style
  - conditional formatting
verified: 2026-11-14
---

Rows and columns are 1-based, as in Excel. `sheet.Cell("B4")` returns a lightweight cell handle and `sheet.Range("A1:D1")` returns a range. Format blocks of cells through a range: one range call applies the style once instead of cell by cell.

## Format cells and ranges

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Budget.xlsx");
ExcelSheet sheet = book.AddSheet("Budget");

sheet.Range("A1:C3").SetValues(new object?[,] {
    { "Item", "Cost", "Share" },
    { "Laptops", 18_500m, 0.72 },
    { "Licenses", 7_250m, 0.28 },
});

sheet.Range("A1:C1")
    .SetBold()
    .SetFillColor(OfficeColor.Parse("D9EAF7"))
    .SetHorizontalAlignment(ExcelHorizontalAlignment.Center)
    .SetBorder(ExcelBorderStyle.Thin, OfficeColor.Parse("8EA9C1"));

sheet.Range("B2:B3").SetNumberFormat("#,##0.00");
sheet.Range("C2:C3").SetNumberFormat("0.0%");
sheet.Cell("A3").SetItalic().SetColor(OfficeColor.Parse("595959"));

sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(24));
sheet.View.Freeze(topRows: 1);

book.Save();
```

### PowerShell

```powershell
New-OfficeExcel -Path .\Budget.xlsx {
    ExcelSheet 'Budget' {
        ExcelCell A1 'Item'
        ExcelCell B1 'Cost'
        ExcelCell A2 'Laptops'
        ExcelCell B2 18500 -NumberFormat '#,##0.00'
        ExcelColumn A -Width 24
    }
}
```

## Format a table column by header

When the data came from objects, style a column by its header name instead of by address. The builder applies number formats such as `Currency`, `Percent`, `Number`, `Date`, and `Integer`.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

var rows = new[] {
    new { Region = "North", Revenue = 27_500m, Margin = 0.21, Status = "Paid" },
    new { Region = "South", Revenue = 8_900m, Margin = 0.17, Status = "Overdue" }
};

using var book = ExcelDocument.Create("Regions.xlsx");
ExcelSheet sheet = book.AddSheet("Regions");
sheet.AddTable(rows, new OfficeTableOptions { Name = "Regions" });

sheet.ColumnStyleByHeader("Revenue").Currency(decimals: 0);
sheet.ColumnStyleByHeader("Margin").Percent(decimals: 1);
sheet.ColumnStyleByHeader("Status").SetFillColorWhenTextEquals("Overdue", OfficeColor.Parse("FEE2E2"));
sheet.AutoFitColumns();

book.Save();
```

### PowerShell

```powershell
Open-OfficeExcel -Path .\Regions.xlsx {
    ExcelSheet 'Regions' {
        ExcelColumn -Header Revenue -Table Regions -Style Currency -Decimals 0
        ExcelColumn -Header Status -Table Regions -BackgroundByText @{ Overdue = '#FEE2E2'; Paid = '#DCFCE7' }
    }
}
```

## Add conditional formatting

Conditional formats stay live in Excel: they re-evaluate when the values change.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Scores.xlsx");
ExcelSheet sheet = book.AddSheet("Scores");
sheet.Range("A1:B4").SetValues(new object?[,] {
    { "Team", "Score" }, { "Blue", 82 }, { "Green", 64 }, { "Red", 91 }
});

sheet.Range("B2:B4").ConditionalFormat().DataBar(OfficeColor.Accent1);
sheet.Range("B2:B4").ConditionalFormat()
    .CellIs(ExcelConditionalFormattingOperator.LessThan, "70", fillColor: OfficeColor.Parse("FEE2E2"));

book.Save();
```

### PowerShell

```powershell
Open-OfficeExcel -Path .\Scores.xlsx {
    ExcelSheet 'Scores' {
        ExcelConditionalFormat -Range B2:B4 -DataBar
    }
}
```
