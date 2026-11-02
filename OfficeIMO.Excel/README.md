# OfficeIMO.Excel

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Excel)](https://www.nuget.org/packages/OfficeIMO.Excel)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Excel?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Excel)

Create, edit, read and save Excel workbooks (`.xlsx`, and the supported subset of Excel 97–2003 `.xls`) in .NET, without Microsoft Excel or COM. Tables, charts and cell styles follow the workbook's theme; very large exports stream rows with bounded memory.

## Returned authoring objects

Conditional-format Add methods return `ExcelConditionalFormattingInfo`, the same
editable model used by `AddConditionalFormattingRule`. Edit the returned model
and call `sheet.UpdateConditionalFormattingRule(rule)` to apply it, or pass it to
`RemoveConditionalFormattingRule` to remove that exact rule. Pivot conditional
formats follow the same contract.

`AddAutoFilter` returns an `ExcelAutoFilter` range handle. Its `Inspect()` method
reads current criteria and table ownership. Creating the handle preserves deferred
workbook exports. `AddTableOfContents` returns the sheet it creates, and
`AddPivotManualGrouping` returns the created `ExcelPivotGroupingInfo`.

`CustomDocumentProperties.Add` returns the attached `ExcelCustomProperty`, so edits
to the returned value are saved with the workbook. The collection also implements
`IDictionary<string, ExcelCustomProperty>` and `ICollection<KeyValuePair<string, ExcelCustomProperty>>`.

`AddManualRowPageBreak` and `AddManualColumnPageBreak` return an attached
`ExcelPageBreak` with its worksheet, index, orientation and `Remove()` method.
`AddMacro` returns the attached `ExcelMacroProject`, which can read, save or remove
the VBA payload. Removed or replaced handles reject further operations.
`AddIgnoredErrorRegion` returns the created `ExcelIgnoredErrorRegionInfo`.
`AddBackLinksToToc` returns hyperlink metadata keyed by the worksheets receiving
links. Use `AddImage` to create an image and receive its `ExcelImage`.
Image dimensions and anchor offsets use `OfficeLength`: `image.Width`, `image.Height`,
`image.OffsetX` and `image.OffsetY` retain whole EMUs. Two-cell ending offsets use
`ToOffsetX` and `ToOffsetY`. For example, replace `widthPixels: 96` with
`new ExcelImageOptions { Width = OfficeLength.Pixels(96) }`. `SetSize`, `MoveTo`, `SetTwoCellEndingMarker` and
`TryGetAbsoluteAnchorBounds` use the same length contract. Cell and ending-marker
offsets on existing XML remain inspectable when signed; `SetTwoCellEndingMarker`
requires non-negative ending offsets; the absolute-position `AddImage(x, y, ...)` overload requires non-negative
authored positions. HTML round trips preserve signed positions and offsets read
from existing native drawings. Omitted byte-image
sizes remain 96 by 32 pixels at 96 DPI, while file/scaled placement retains its
source-dimension defaults. Remote insertion also uses detected source dimensions
when width, height and scale are omitted; use explicit lengths to request a fixed
96 by 32 pixel image. Raster export applies its pixel rounding and safety
limits when rendering. `GetRangeSize` returns `Width` and `Height` as lengths;
worksheet reports insert images through `sheet.Report().AddImage(...)`.


Use `AddImage` with a path, stream or native byte array. `AddImageAsync` reads files
and streams asynchronously or loads an HTTP or HTTPS `Uri`. Geometry, metadata,
network policy and stream ownership belong to `ExcelImageOptions`:

```csharp
ExcelImage logo = sheet.AddImage(2, 3, "logo.png", new ExcelImageOptions {
    ScalePercent = 20,
    Description = "Company logo",
    Title = "Logo"
});
sheet.AddImage("A1:C15", File.ReadAllBytes("header.png"), new ExcelImageOptions {
    Placement = ExcelImagePlacement.MoveAndSize
});
await sheet.AddImageAsync(4, 2, new Uri("https://example.org/logo.png"));
```

A single A1 reference uses a one-cell anchor; a colon-delimited range, including
`A1:A1`, uses a two-cell anchor. `AddImage(x, y, source, options)` uses an absolute
anchor. Range anchors determine their own size and reject explicit dimensions or
scale. `LockAspectRatio` sets Excel's picture lock; it does not calculate an omitted
dimension. Streams stay open by default and retain their seekable position;
`LeaveOpen = false` closes them on success, failure or cancellation. Image bytes
are embedded without re-encoding. `Deduplicate = true` shares identical native
payloads within the worksheet drawing while each picture keeps its own metadata;
the default creates separate image parts. `sheet.Report()` uses the same options
and returns the worksheet's attached image.


## Install

```sh
dotnet add package OfficeIMO.Excel
dotnet add package OfficeIMO.Excel.Pdf   # only for SaveAsPdf
```

Targets `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.

## A table and a chart from your data

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

`AddTable` writes a native Excel table: a header row from the property names, typed cells and the theme's table style. It returns the `ExcelTable`, and `AddChart` returns the `ExcelChart`, for further changes. `ExcelChartOptions.Row` and `Column` anchor the chart; `Width` and `Height` take an `OfficeLength`.

## Cells, formulas and formatting

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Budget.xlsx");
ExcelSheet sheet = book.AddSheet("Budget");

sheet.CellValue(1, 1, "Item");
sheet.CellValue(1, 2, "Cost");
sheet.CellValue(2, 1, "Laptops");
sheet.CellValue(2, 2, 18_500m);
sheet.CellValue(3, 1, "Licenses");
sheet.CellValue(3, 2, 7_250m);
sheet.CellValue(4, 1, "Total");
sheet.CellFormula(4, 2, "SUM(B2:B3)");

sheet.FormatCell(2, 2, "#,##0.00");
sheet.FormatCell(3, 2, "#,##0.00");
sheet.FormatCell(4, 2, "#,##0.00");
sheet.AddConditionalDataBar("B2:B3", OfficeColor.Accent1);
sheet.Freeze(topRows: 1);
sheet.AutoFitColumns();

book.Save();
```

Rows and columns are 1-based, as in Excel.

## Read rows back as objects

```csharp
using OfficeIMO.Excel;

using var book = ExcelDocument.Load("Orders.xlsx");
foreach (Order order in book["Orders"].RowsAs<Order>()) {
    Console.WriteLine($"{order.Customer}: {order.Total:N0}");
}

record Order(string Customer, string Region, decimal Total, double Margin);
```

Columns map to properties or constructor parameters by header name. `RowsAs` also accepts a factory or an explicit column mapping. For files from outside your organization, load with `ExcelLoadOptions.UntrustedDefaults`.

XML Spreadsheet import also limits defined names (10,000 by default). Set `ExcelSpreadsheetXmlOptions.MaxNames` for a larger trusted workbook.

## Large exports

The object model keeps the whole workbook editable until `Save()`. For hundreds of thousands of rows, the streaming writer writes each row as it arrives and keeps memory flat:

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

IEnumerable<LogEvent> events = Enumerable.Range(1, 200_000)
    .Select(i => new LogEvent(DateTime.UtcNow.AddSeconds(-i), i % 7 == 0 ? "Warning" : "Info", $"Event {i}"));

using (ExcelWorkbookWriter writer = ExcelDocument.CreateWriter("Events.xlsx", theme: OfficeTheme.Modern)) {
    writer.WriteTable("Events", events);
}

record LogEvent(DateTime Time, string Level, string Message);
```

`WriteTable` also takes an `IAsyncEnumerable<T>`, a `DataTable` or an `IDataReader`. The file is published when the writer is disposed.

## Report sheets

`sheet.Report()` lays out titled sections, KPI rows, paragraphs and tables on a sheet, for workbooks that people read rather than filter:

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

using var book = ExcelDocument.Create("Summary.xlsx");
book.AddSheet("Summary").Report()
    .Title("Service review", "October")
    .KpiRow([("Availability", "99.95%"), ("Incidents", 3), ("Changes", 41)])
    .Section("Highlights")
    .BulletedList(["No priority 1 incidents", "Patch compliance at 98%"]);
book.Save();
```

## Save as PDF

```csharp
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;

using var book = ExcelDocument.Load("Orders.xlsx");
OfficeConversionReport report = book.SaveAsPdf("Orders.pdf");
Console.WriteLine($"{report.ItemsWritten} pages, loss: {report.HasLoss}");
```

## More

- [Excel guide](../Docs/officeimo.excel-guide.md): styles, validation, conditional formats, pivot tables, slicers, sparklines, templates, formulas and calculation, protection, links, images, printing, legacy `.xls` and image export.
- [Formula calculation](../Docs/officeimo.excel-guide.md#formula-inspection-and-calculation-policy): typed errors and lookup modes, reference-valued `INDEX` and `XLOOKUP`, and named scalar and array expressions.
- [Compatibility matrix](COMPATIBILITY.md) and the [generated operation contract](../Docs/Compatibility/generated/package-operations.md): exactly what is supported, preserved or rejected.
- [Large-workbook guidance](../Docs/officeimo.excel.large-workbook-guidance.md), [remote loading](../Docs/officeimo.excel.remote-loading.md) and [legacy XLS compatibility](../Docs/officeimo.excel.legacy-xls-compatibility.md).
- [Examples](../OfficeIMO.Examples/Excel) and the [migration guide](../MIGRATION.md).
- Adapters: [OfficeIMO.Excel.Pdf](../OfficeIMO.Excel.Pdf/README.md), [Html](../OfficeIMO.Excel.Html/README.md), [Csv](../OfficeIMO.Excel.Csv/README.md), [OpenDocument](../OfficeIMO.Excel.OpenDocument/README.md), [GoogleSheets](../OfficeIMO.Excel.GoogleSheets/README.md) and [IWork](../OfficeIMO.Excel.IWork/README.md). PDF tables back to Excel live in [OfficeIMO.Pdf.Excel](../OfficeIMO.Pdf.Excel/README.md).
- PowerShell: [PSWriteOffice](../PowerShell/PSWriteOffice/README.MD) (`Export-OfficeExcel`, `New-OfficeExcel`).

Dependencies: the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and `OfficeIMO.Core`. MIT licensed.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 22 | 13 | 0 | 15 | 1 | 2 |
| Read | 34 | 0 | 17 | 0 | 0 | 2 |
| Edit | 11 | 0 | 39 | 2 | 1 | 0 |
| Preserve | 10 | 0 | 39 | 0 | 0 | 0 |
| Inspect | 6 | 0 | 0 | 0 | 0 | 0 |
| Validate | 4 | 0 | 0 | 0 | 0 | 2 |
| Remove | 4 | 0 | 0 | 0 | 1 | 0 |
| Convert | 49 | 40 | 0 | 8 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Excel` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
