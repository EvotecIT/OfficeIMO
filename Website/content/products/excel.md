---
title: "OfficeIMO.Excel"
description: "Read and write Excel workbooks with tables, charts, and formatting. No Excel installation required."
layout: product
product_color: "#059669"
install: "dotnet add package OfficeIMO.Excel"
nuget: "OfficeIMO.Excel"
docs_url: "/docs/excel/"
api_url: "/api/excel/"
---

## Why OfficeIMO.Excel?

OfficeIMO.Excel lets you build and consume `.xlsx` workbooks entirely in managed code. Generate dashboards, data exports, financial models, and bulk reports without ever touching COM or requiring an Office license. The API is designed for real-world scenarios -- from one-off scripts to high-throughput server pipelines.

## Features

- **Worksheets & cell values** -- strings, numbers, dates, booleans, and shared strings with full type fidelity
- **Tables with AutoFilter** -- structured tables with column headers, totals row, and built-in filter controls
- **Named ranges & formulas** -- workbook and sheet-scoped names, cell formulas, and calculated columns
- **Charts** -- column, pie, doughnut, scatter, and bubble charts with series data, axis labels, and legends
- **Conditional formatting** -- color scales, data bars, icon sets, and rule-based highlight formatting
- **Validation** -- list, whole number, decimal, date, time, text length, and custom formula validators
- **Pivot tables & sparklines** -- summarize large data sets and embed inline sparklines in cells
- **Parallel execution** -- bulk read/write operations optimized for multi-core workloads
- **Images & hyperlinks** -- embed images in cells and attach hyperlinks to cells or shapes
- **AutoFit columns** -- automatically size columns to fit content width
- **Headers, footers & print setup** -- page headers, footers, margins, orientation, and print area

## Quick start

```csharp
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Create("Sales.xlsx");
var sheet = workbook.AddSheet("Q4 Sales");

// Set headers
sheet.Cells["A1"].Value = "Product";
sheet.Cells["B1"].Value = "Units";
sheet.Cells["C1"].Value = "Revenue";

// Add data rows
string[] products = { "Widget A", "Widget B", "Widget C", "Widget D" };
int[] units = { 1200, 850, 2100, 430 };
decimal[] revenue = { 24000m, 17000m, 63000m, 12900m };

for (int i = 0; i < products.Length; i++)
{
    int row = i + 2;
    sheet.Cells[$"A{row}"].Value = products[i];
    sheet.Cells[$"B{row}"].Value = units[i];
    sheet.Cells[$"C{row}"].Value = revenue[i];
    sheet.Cells[$"C{row}"].NumberFormat = "$#,##0";
}

// Create a table from the data range
var table = sheet.AddTable("SalesTable", "A1", $"C{products.Length + 1}");
table.Style = ExcelTableStyle.Medium9;
table.ShowTotalsRow = true;
table.Columns["Revenue"].TotalsFunction = ExcelTotalsFunction.Sum;

// AutoFit for a clean layout
sheet.AutoFitColumns();

workbook.Save();
```

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.Excel runs on Windows, Linux, and macOS. Generated workbooks are fully compatible with Microsoft Excel, LibreOffice Calc, and Google Sheets.
