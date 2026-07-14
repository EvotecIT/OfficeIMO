# OfficeIMO.Excel - Excel workbooks for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Excel)](https://www.nuget.org/packages/OfficeIMO.Excel)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Excel?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Excel)

`OfficeIMO.Excel` is the main Excel package in the OfficeIMO family. It creates, edits, reads, converts, and saves `.xlsx` workbooks without COM automation and without Microsoft Excel installed. It can also open BIFF8 legacy binary `.xls` workbooks, project supported content into the normal OfficeIMO Excel model, and save a native `.xls` subset through the first-party BIFF writer.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Excel
```

## Quick start

```csharp
using OfficeIMO.Excel;

using var document = ExcelDocument.Create("report.xlsx");
var sheet = document.AddWorksheet("Data");

sheet.CellValue(1, 1, "Name");
sheet.CellValue(1, 2, "Value");
sheet.CellValue(2, 1, "Alpha");
sheet.CellValue(2, 2, 42);
sheet.AddTable("A1:B2", hasHeader: true, name: "DataTable", style: TableStyle.TableStyleMedium9);
sheet.AutoFitColumns();

document.Save();
```

## What it does

- Creates and edits workbooks, worksheets, cells, ranges, tables, styles, hyperlinks, formulas, names, comments, images, charts, filters, and page setup.
- Reads values through worksheet, range, row, dictionary, stream, and typed object helpers.
- Supports editable row workflows where rows can be read, changed, and saved back.
- Handles practical workbook hygiene such as table/filter conflicts, safe table names, deterministic save order, and feature inspection.
- Includes parallel execution controls for heavy export and autofit workloads while serializing the Open XML mutation phase safely.

## Competitive performance with workbook features

OfficeIMO.Excel is optimized for fast tabular reads and writes, but it is not
only a streaming data pipe. The same first-party model authors and edits styles,
tables, formulas, charts, pivots, conditional formatting, validation, images,
templates, protection, print settings, headers and footers, and both `.xlsx`
and the supported legacy `.xls` subset.

The compact table deliberately mixes raw data paths with feature-bearing work:
typed-object reads, plain and styled `DataReader` exports, and a report containing
normal workbook features. Each row only includes libraries with a directly
comparable public API. Lower is faster.
Differences below 5% are treated as ties rather than ranking claims.

<!-- officeimo-excel-benchmark-table:start -->
| Scenario | Variables | Host | Operation | Metric | OfficeIMO.Excel | ClosedXML | EPPlus | LargeXlsx | SpreadCheetah | Sylvan.Data.Excel | Result |
| --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| Compact DataReader to XLSX | Format=.xlsx, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14 | .NET 8 | Write | MeanMs | 1.00x (23ms) | n/a | n/a | 1.14x (27ms) | 1.05x (25ms) | 1.16x (27ms) | OfficeIMO.Excel fastest |
| Feature-rich report to XLSX | Format=.xlsx, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14 | .NET 8 | Create | MeanMs | 1.00x (38ms) | n/a | 9.31x (355ms) | n/a | n/a | n/a | OfficeIMO.Excel fastest |
| Styled DataReader table to XLSX | Format=.xlsx, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14 | .NET 8 | Write | MeanMs | 1.00x (35ms) | 8.68x (301ms) | 8.10x (281ms) | n/a | n/a | n/a | OfficeIMO.Excel fastest |
| Typed objects streamed from XLSX | Format=.xlsx, Rows=25,000, Runner=rotated local, Snapshot=2026-07-14 | .NET 8 | Read | MeanMs | 1.00x (56ms) | 4.84x (272ms) | 3.80x (213ms) | n/a | n/a | 0.67x (38ms) | OfficeIMO.Excel slower than Sylvan.Data.Excel |
<!-- officeimo-excel-benchmark-table:end -->

These are local direction-finding results, not guarantees. Hardware, runtime,
workload shape, package versions, warm-up, and library options change outcomes;
results will vary. OfficeIMO wins some lanes and not others. The
[benchmark harness](../OfficeIMO.Excel.Benchmarks/README.md) publishes the full
comparison suite against ClosedXML, EPPlus, MiniExcel, LargeXlsx,
SpreadCheetah, ExcelDataReader, and Sylvan.Data.Excel. The opt-in
[NPOI comparison](../OfficeIMO.Excel.Benchmarks.NPOI/README.md) separately covers
`.xlsx` row/cell work and legacy `.xls` values, formulas, metadata, formatting,
filters, styles, and pictures.

## Examples

The quick start covers the smallest workbook. These examples show common read, write, reporting, and automation workflows that belong in `OfficeIMO.Excel`.

### Read rows by header

```csharp
using var document = ExcelDocument.Load("input.xlsx");
var sheet = document["Data"];

foreach (var row in sheet.Rows()) {
    Console.WriteLine(row["Name"]);
}
```

### Work with legacy XLS workbooks

```csharp
using var document = ExcelDocument.Load("legacy.xls");
ExcelFeatureReport report = document.InspectFeatures();

document.Save("converted.xlsx");
document.Save("native-copy.xls");

ExcelDocument.Convert("legacy.xls", "converted.xlsx");
ExcelDocument.Convert("openxml.xlsx", "native-copy.xls");
```

BIFF8 `.xls` files load through the normal `ExcelDocument.Load` entry point.
Supported cells, formulas, styles, names, comments, filters, validations,
conditional formatting, layout, protection metadata, document properties,
images, drawings, tables, and chart sheets project into the normal OfficeIMO
model. Unsupported sheet kinds, VBA, embedded OLE content, signatures, and
unprojected BIFF records are reported through the legacy import diagnostics
instead of being silently dropped.

Native `.xls` save uses the same `Save("*.xls")` path as other OfficeIMO saves.
When a workbook contains a feature outside the supported BIFF8 writer subset,
OfficeIMO throws a preflight error with the unsupported feature name so the
caller can save as `.xlsx`, remove the feature, or choose a different workflow.
`ExcelDocument.Convert(...)` uses those same load and save paths and blocks
legacy sources with unsupported or preserve-only content by default. Set
`LossPolicy` to `ExcelConversionLossPolicy.Allow` on conversion or save options
only after reviewing that loss. See
[XLS and XLSX compatibility](../Docs/officeimo.excel.legacy-xls-roadmap.md) for
the current capability matrix, safety contract, and breaking API migration.

### Map rows to objects

```csharp
using var document = ExcelDocument.Load("input.xlsx");
List<Person> people = document["Data"].RowsAs<Person>("A1:C100").ToList();

public sealed class Person {
    public string Name { get; set; } = "";
    public int Value { get; set; }
    public string Status { get; set; } = "";
}
```

### Append to an existing table

```csharp
using var document = ExcelDocument.Load("sales.xlsx");
var rows = new DataTable();
rows.Columns.Add("Revenue", typeof(decimal));
rows.Columns.Add("Region", typeof(string));
rows.Rows.Add(150m, "APAC");

document["Sales"].AppendDataTableToTable(rows, "SalesTable");
document.Save();
```

### Validation lists and typed reads

```csharp
using var document = ExcelDocument.Load("input.xlsx");
var sheet = document["Data"];

sheet.ValidationList("C2:C100", new[] { "New", "Processed", "Hold" });
sheet.Range("D2:D100").Validate.WholeNumberBetween(1, 10, errorMessage: "Use 1 through 10");

List<RowModel> rows = document.Read()
    .Sheet("Data")
    .Range("A1:C100")
    .AsObjects<RowModel>()
    .ToList();

public sealed class RowModel {
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}
```

### Charts and dashboard recipes

```csharp
using OfficeIMO.Excel;

using var document = ExcelDocument.Create("dashboard.xlsx");
var sheet = document.AddWorksheet("Summary");

sheet.CellValue(1, 1, "Quarter");
sheet.CellValue(1, 2, "Revenue");
sheet.CellValue(2, 1, "Q1");
sheet.CellValue(2, 2, 10);
sheet.CellValue(3, 1, "Q2");
sheet.CellValue(3, 2, 18);
sheet.CellValue(4, 1, "Q3");
sheet.CellValue(4, 2, 24);
sheet.CellValue(5, 1, "Q4");
sheet.CellValue(5, 2, 30);

sheet.AddTable("A1:B5", hasHeader: true, name: "RevenueTable", style: TableStyle.TableStyleMedium2);
sheet.ChartFromTable("RevenueTable")
    .RevenueTrend("Revenue trend")
    .Size(640, 320)
    .At(row: 1, column: 5);

document.Save();
```

### Pivot tables and pivot-backed charts

```csharp
using OfficeIMO.Excel;
using System.Linq;

using var document = ExcelDocument.Create("pivot-report.xlsx");
var sheet = document.AddWorksheet("Sales");

sheet.CellValue(1, 1, "Region");
sheet.CellValue(1, 2, "Product");
sheet.CellValue(1, 3, "Quarter");
sheet.CellValue(1, 4, "Revenue");
sheet.CellValue(2, 1, "EMEA");
sheet.CellValue(2, 2, "Alpha");
sheet.CellValue(2, 3, "Q1");
sheet.CellValue(2, 4, 125000);
sheet.CellValue(3, 1, "EMEA");
sheet.CellValue(3, 2, "Beta");
sheet.CellValue(3, 3, "Q1");
sheet.CellValue(3, 4, 94000);
sheet.CellValue(4, 1, "APAC");
sheet.CellValue(4, 2, "Alpha");
sheet.CellValue(4, 3, "Q2");
sheet.CellValue(4, 4, 141000);
sheet.AddTable("A1:D4", hasHeader: true, name: "SalesTable", style: TableStyle.TableStyleMedium4);

sheet.Pivot("A1:D4")
    .Rows("Region")
    .Columns("Quarter")
    .Filters("Product")
    .Sum("Revenue", "Total revenue", "#,##0")
    .Layout(ExcelPivotLayout.Tabular)
    .Style("PivotStyleMedium9")
    .Captions(rowHeader: "Region", columnHeader: "Quarter", grandTotal: "Total")
    .At("F2", "SalesPivot");

var pivot = sheet.GetPivotTables().Single(p => p.Name == "SalesPivot");
Console.WriteLine($"{pivot.Name}: {string.Join(", ", pivot.RowFields)}");

var chart = sheet.ChartFromTable("SalesTable")
    .VarianceColumns("Revenue by region")
    .At(row: 12, column: 1);
chart.SetPivotSource("SalesPivot");

document.Save();
```

Pivot support is useful but still marked partial in the compatibility matrix. It covers source-range pivots, row/column/page/data fields, styles, layouts, filters, calculated fields, grouping metadata, and readback. Slicers, timelines, external connections, and query-table authoring are still preserve-oriented or roadmap areas.

### Formula inspection and calculation policy

```csharp
using var document = ExcelDocument.Load("report.xlsx");

var formulas = document.InspectFormulas();
Console.WriteLine(formulas.ToMarkdown());

foreach (var formula in formulas.Formulas.Where(f => !f.IsSupportedByOfficeIMO)) {
    Console.WriteLine($"{formula.SheetName}!{formula.CellReference}: {formula.UnsupportedReason}");
}

int calculated = document.Calculate();
document.Save("report.xlsx", new ExcelSaveOptions {
    EvaluateFormulasBeforeSave = true,
    ForceFullCalculationOnOpen = true
});
```

### Preflight a workbook before choosing a workflow

```csharp
using var document = ExcelDocument.Load("incoming.xlsx", readOnly: true);

ExcelFeatureReport report = document.InspectFeatures();

try {
    report.EnsureCan(ExcelPreflightCapability.EditWorkbookStructure);
} catch (InvalidOperationException ex) {
    Console.WriteLine(ex.Message);
}

if (!report.Can(ExcelPreflightCapability.ExportPdfReport)) {
    Console.WriteLine(report.ToMarkdown());
}
```

Use workflow preflight when an application needs to decide whether a workbook is safe for readback, cell-value edits, structure-changing edits, cached-formula reads, OfficeIMO formula calculation, template binding, or first-party PDF report export. Preserve-only features such as macros, slicers, timelines, threaded comments, external links, custom XML, OLE objects, and form controls are reported with package details instead of being silently ignored.

### DataTable and JSON exchange

```csharp
using System.Data;

using var document = ExcelDocument.Load("data.xlsx");
var sheet = document["Data"];

DataTable table = sheet.ToDataTable("A1:C100");
string json = sheet.ToJson("A1:C100");

sheet.FromJson("[{\"Name\":\"Gamma\",\"Amount\":30}]", startRow: 8, startColumn: 1);
```

### Template markers

```csharp
using var document = ExcelDocument.Load("invoice-template.xlsx");

int replacements = document.ApplyTemplate(new Dictionary<string, object?> {
    ["Invoice.Number"] = "INV-001",
    ["Customer.Name"] = "Adatum",
    ["Total"] = 123.45m
});

var template = document.InspectTemplate(new {
    Invoice = new { Number = "INV-001" },
    Customer = new { Name = "Adatum" },
    Total = 123.45m
});

template.EnsureAllMarkersBound();
document.Save("invoice.xlsx");
```

### Comments and conditional formatting

```csharp
using var document = ExcelDocument.Load("review.xlsx");
var sheet = document["Data"];

sheet.SetComment("A1", "Review total", author: "Alice", initials: "AA");
sheet.UpdateComments(new ExcelCommentFilter { TextContains = "total" }, "Total reviewed", author: "Carol", initials: "CC");

sheet.AddConditionalColorScale("C2:C100", "#FFF0F0", "#70AD47");
sheet.Range("D2:D100").ConditionalFormat.DataBar("#5B9BD5");

document.Save();
```

### Tune larger exports

```csharp
using var document = ExcelDocument.Create("large-report.xlsx");
document.Execution.Mode = ExecutionMode.Automatic;
document.Execution.MaxDegreeOfParallelism = Environment.ProcessorCount;
document.Execution.SaveWorksheetAfterAutoFit = false;
```

For a new workbook that only contains tabular data, write the XLSX package
directly without building an editable workbook model:

```csharp
using var output = File.Create("large-export.xlsx");

ExcelDocument.WriteRows(
    output,
    rows,
    new[] { "Id", "Name", "Created", "Active" },
    static (writer, row) => writer
        .Write(row.Id)
        .Write(row.Name)
        .Write(row.Created)
        .Write(row.Active),
    new ExcelTabularWriteOptions {
        SheetName = "Data",
        IncludeCellReferences = false,
        UseSharedStrings = false
    });
```

### Fluent compose

```csharp
using var document = ExcelDocument.Create("composed-report.xlsx");

document.Compose("Report", composer => {
    composer.Title("Demo Report", "Generated with OfficeIMO.Excel");
    composer.Callout("info", "Heads up", "Generated via the fluent API");
    composer.Section("Summary");
    composer.PropertiesGrid(new (string, object?)[] {
        ("Author", "OfficeIMO"),
        ("Date", DateTime.Today.ToString("yyyy-MM-dd"))
    });

    var items = new[] {
        new { Name = "Alice", Score = 90, Status = "OK" },
        new { Name = "Bob", Score = 80, Status = "Warning" }
    };

    composer.TableFrom(items, title: "Scores", visuals: visuals => {
        visuals.NumericColumnDecimals["Score"] = 0;
        visuals.TextBackgrounds["Status"] = new Dictionary<string, string> {
            ["Warning"] = "#FFF3CD"
        };
    });

    composer.HeaderFooter(header => header.Center("Demo Report").FooterRight("Page &P of &N"));
    composer.Finish(autoFitColumns: true);
});

document.Save();
```

## Adjacent packages

| Package | Use it for |
| --- | --- |
| [OfficeIMO.Excel.Pdf](../OfficeIMO.Excel.Pdf/README.md) | Excel to PDF export through `OfficeIMO.Pdf`, plus PDF table import to Excel. |
| [OfficeIMO.Excel.GoogleSheets](../OfficeIMO.Excel.GoogleSheets/README.md) | Planning and exporting Excel content to Google Sheets. |
| [OfficeIMO.Excel.Benchmarks](../OfficeIMO.Excel.Benchmarks/README.md) | Benchmark harness for Excel workloads. |

## Deeper docs

- [Compatibility matrix](COMPATIBILITY.md)
- [Large workbook guidance](../Docs/officeimo.excel.large-workbook-guidance.md)
- [Excel roadmap](../Docs/officeimo.excel.roadmap.md)
- [Excel examples](../OfficeIMO.Examples/Excel)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** Open XML SDK for `.xlsx` package mechanics. Microsoft BCL/JSON compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Drawing`. The workbook API, BIFF8 `.xls` reader/writer, large-data paths, validation, and PNG/SVG export are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
