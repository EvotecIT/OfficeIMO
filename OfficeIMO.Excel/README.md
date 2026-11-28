# OfficeIMO.Excel

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Excel)](https://www.nuget.org/packages/OfficeIMO.Excel)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Excel?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Excel)

Create, edit, read and convert Excel workbooks in .NET without Microsoft Excel or COM. `OfficeIMO.Excel` works with `.xlsx`, `.xlsm`, templates and add-ins, and the supported subsets of `.xlsb` and Excel 97–2003 `.xls`. Tables, charts and cell styles follow the workbook's theme, and very large exports stream rows with bounded memory.

## Install

```sh
dotnet add package OfficeIMO.Excel
```

Add [OfficeIMO.Excel.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.Pdf/README.md) only if you need `SaveAsPdf`.

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/Excel/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO;
using OfficeIMO.Excel;

Order[] orders = [
    new("Contoso", "North", 12_400m, 0.21),
    new("Fabrikam", "South", 8_900m, 0.17),
    new("Northwind", "North", 15_100m, 0.24),
    new("Tailspin", "West", 6_300m, 0.12)
];

using var book = ExcelDocument.Create("Orders.xlsx", new ExcelCreateOptions { Theme = OfficeTheme.Modern });
ExcelSheet sheet = book.AddSheet("Orders");

ExcelTable table = sheet.AddTable(orders, new OfficeTableOptions { Name = "Orders" });
sheet.AddChart(OfficeChartType.Column, orders, nameof(Order.Customer),
    new ExcelChartOptions { Title = "Revenue by customer", Row = 2, Column = 7 }, nameof(Order.Total));

book.Save();

record Order(string Customer, string Region, decimal Total, double Margin);
```
<!-- /powerforge:code -->

`AddTable` writes a native Excel table with a header row from the property names, typed cells and the theme's table style. `AddChart` returns the `ExcelChart` for further changes.

## What it does

- **Write cells, ranges and formulas** with number formats, fonts, fills, borders, conditional formats, data validation, freeze panes and AutoFit. Rows and columns are 1-based, as in Excel.
- **Calculate formulas** with a built-in engine, including typed errors, lookups, reference-valued `INDEX` and `XLOOKUP`, and named scalar and array expressions, and keep the cached results Excel shows.
- **Tables, pivot tables, slicers, sparklines and charts** as native, editable Excel objects.
- **Read rows back as objects** with `RowsAs<T>`, or stream a sheet forward-only through `ExcelDocument.OpenReader`, a `DbDataReader`.
- **Stream large exports** with `ExcelDocument.CreateWriter`, which writes XLSX, XLSB or XLS row by row from an `IEnumerable<T>`, `IAsyncEnumerable<T>`, `DataTable` or `IDataReader`.
- **Report sheets** with `sheet.Report()`: titles, KPI rows, sections, lists and tables for workbooks people read rather than filter.
- **Templates, protection, links, images, printing and VBA projects**, preserved and editable.
- **Before sharing**: `InspectHiddenData` and `RemoveHiddenData` find and remove authors, comments, hidden sheets, rows and names, external links and image metadata.
- **Render** charts and sheets to PNG, SVG and other image formats.
- **Load untrusted files** with `ExcelLoadOptions.UntrustedDefaults`; XLSB and package limits are checked before a workbook is opened.

The [Excel guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.excel-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| PDF | [OfficeIMO.Excel.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.Pdf/README.md) |
| HTML tables, both directions | [OfficeIMO.Excel.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.Html/README.md) |
| CSV, both directions | [OfficeIMO.Excel.Csv](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.Csv/README.md) |
| OpenDocument spreadsheets | [OfficeIMO.Excel.OpenDocument](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.OpenDocument/README.md) |
| Google Sheets | [OfficeIMO.Excel.GoogleSheets](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.GoogleSheets/README.md) |
| Apple Numbers | [OfficeIMO.Excel.IWork](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.IWork/README.md) |
| PDF tables back to Excel | [OfficeIMO.Pdf.Excel](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Pdf.Excel/README.md) |

Saving to another format returns an `OfficeConversionReport` that lists anything the target couldn't carry.

## Limits

`.xlsx` is the native format. `.xlsb` and `.xls` load into the same model and save natively within their documented subsets; binary streaming writers create plain worksheets. The [compatibility matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel/COMPATIBILITY.md), [legacy XLS compatibility](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.excel.legacy-xls-compatibility.md), [large-workbook guidance](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.excel.large-workbook-guidance.md) and the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md) list exactly what is supported, preserved or rejected.

## Learn more

- [OfficeIMO.Excel on officeimo.com](https://officeimo.com/products/excel/) and the [Excel docs](https://officeimo.com/docs/excel/).
- [Excel guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.excel-guide.md), including [formula calculation](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.excel-guide.md#formula-inspection-and-calculation-policy) and [remote loading](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.excel.remote-loading.md).
- [Examples](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Excel) and the [migration guide](https://github.com/EvotecIT/OfficeIMO/blob/master/MIGRATION.md).
- PowerShell: [PSWriteOffice](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/README.MD) (`Export-OfficeExcel`, `New-OfficeExcel`). Command line: `officeimo convert book.xlsx book.pdf` with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and `OfficeIMO.Core`. Add `OfficeIMO.Security` only to validate package signatures.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 25 | 19 | 0 | 13 | 1 | 2 |
| Read | 43 | 0 | 15 | 0 | 0 | 2 |
| Edit | 12 | 0 | 45 | 2 | 1 | 0 |
| Preserve | 11 | 0 | 45 | 0 | 0 | 0 |
| Inspect | 6 | 0 | 0 | 0 | 0 | 0 |
| Validate | 4 | 0 | 0 | 0 | 0 | 2 |
| Remove | 4 | 0 | 0 | 0 | 1 | 0 |
| Convert | 61 | 43 | 0 | 8 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Excel` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
