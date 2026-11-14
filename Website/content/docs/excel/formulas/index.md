---
title: Formulas and Named Ranges
description: Write Excel formulas, define named ranges, calculate supported functions in .NET, and inventory formulas in existing workbooks with OfficeIMO.Excel.
order: 25
packages:
  - OfficeIMO.Excel
limits: "Formula calculation covers the supported function set; other functions are written unchanged and calculated by the spreadsheet application."
---

Formulas are stored as Excel formulas, so Excel and other spreadsheet applications recalculate them on open. OfficeIMO.Excel can also calculate the functions it supports and cache the results, so readers that do not recalculate (search indexers, previews, `GetValue`) see the numbers.

## Write formulas

Write the formula text without the leading `=`.

### C#

```csharp
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Invoice.xlsx");
ExcelSheet sheet = book.AddSheet("Invoice");

sheet.Range("A1:C4").SetValues(new object?[,] {
    { "Item", "Quantity", "Price" },
    { "Cable", 4, 12.50m },
    { "Switch", 1, 189.00m },
    { "Rack", 1, 420.00m },
});
sheet.Cell("D1").SetValue("Line total");
for (int row = 2; row <= 4; row++) {
    sheet.Cell(row, 4).SetFormula($"B{row}*C{row}");
}
sheet.Cell("C6").SetValue("Total");
sheet.Cell("D6").SetFormula("SUM(D2:D4)");
sheet.Range("C2:D6").SetNumberFormat("#,##0.00");

int calculated = book.Calculate();
Console.WriteLine($"{calculated} formulas calculated, total = {sheet.Cell("D6").GetValue()}");

book.Save();
```

When a workbook will be opened by Excel and you want every formula recalculated there, call `book.ConfigureFullCalculationOnOpen()` before saving.

### PowerShell

```powershell
New-OfficeExcel -Path .\Estimate.xlsx {
    ExcelSheet 'Estimate' {
        ExcelCell A1 'Servers'
        ExcelCell B1 42
        ExcelCell A2 'Cost per server'
        ExcelCell B2 1850 -NumberFormat '#,##0.00'
        ExcelCell A3 'Total'
        ExcelCell B3 -Formula 'B1*B2' -NumberFormat '#,##0.00'
    }
}
```

## Use named ranges

A defined name keeps formulas readable and survives rows being inserted above the referenced cell.

### C#

```csharp
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Pricing.xlsx");
ExcelSheet inputs = book.AddSheet("Inputs");
inputs.Cell("A1").SetValue("VAT rate");
inputs.Cell("B1").SetValue(0.23);
book.SetNamedRange("VatRate", "'Inputs'!$B$1");

ExcelSheet prices = book.AddSheet("Prices");
prices.Cell("A1").SetValue(100m);
prices.Cell("B1").SetFormula("A1*(1+VatRate)");

book.Save();
```

### PowerShell

```powershell
Open-OfficeExcel -Path .\Pricing.xlsx {
    ExcelSheet 'Inputs' {
        ExcelNamedRange VatRate 'B1' -Global
    }
}
Get-OfficeExcelNamedRange -Path .\Pricing.xlsx
```

## Inventory formulas in existing workbooks

Before changing a model, list its formulas, including volatile functions and external references.

### PowerShell

```powershell
Get-OfficeExcelFormula -Path .\Models\*.xlsx -Volatile | Format-Table SheetName, Address, Formula
Get-ChildItem -LiteralPath .\Models -Filter *.xlsx | Get-OfficeExcelFormula -External
```
