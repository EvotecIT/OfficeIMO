# OfficeIMO.Excel - Excel workbooks for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Excel)](https://www.nuget.org/packages/OfficeIMO.Excel)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Excel?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Excel)

`OfficeIMO.Excel` is the main Excel package in the OfficeIMO family. It creates, edits, reads, converts, and saves `.xlsx` workbooks without COM automation and without Microsoft Excel installed. It also opens BIFF8 `.xls` and BIFF12 `.xlsb` workbooks, projects supported content into the normal OfficeIMO model, and provides first-party native writer subsets with explicit preservation and loss diagnostics.

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

`AsFluent()` wraps the same `ExcelDocument`; it does not create a separate
workbook model. Call `End()` when direct worksheet APIs are more convenient:

```csharp
using var document = ExcelDocument.Create("report.xlsx");

document.AsFluent()
    .Sheet("Data", sheet => sheet
        .Cell(1, 1, "Name")
        .Cell(1, 2, "Value")
        .Cell(2, 1, "Alpha")
        .Cell(2, 2, 42))
    .End();

document.Save();
```

## What it does

- Creates and edits workbooks, worksheets, cells, ranges, tables, styles, hyperlinks, formulas, names, comments, images, charts, filters, and page setup.
- Reads tabular values through the forward-only `ExcelDocument.OpenDataReader(...)` API and typed `ExcelSheet.RowsAs<T>(...)` helpers.
- Edits loaded workbooks through the normal worksheet, cell, range, table, and fluent authoring APIs.
- Handles practical workbook hygiene such as table/filter conflicts, safe table names, deterministic save order, and feature inspection.
- Applies optional shared package-security policy before parsing Open XML, XLSB, or compound XLS files.
- Includes parallel execution controls for heavy export and autofit workloads while serializing the Open XML mutation phase safely.

## Performance evidence

OfficeIMO.Excel is optimized for fast tabular reads and writes, but it is not
only a streaming data pipe. The same first-party model authors and edits styles,
tables, formulas, charts, pivots, conditional formatting, validation, images,
templates, protection, print settings, headers and footers, and both `.xlsx`
and the supported legacy `.xls` subset.

Performance claims use validated outputs and record the workload, package versions, runtime, operating system, processor, warm-up, iterations, allocations, and source provenance. Windows, Linux, and macOS remain separate evidence lanes; missing platforms stay visible rather than being inferred from another operating system.

Use the [benchmark website](https://officeimo.com/benchmarks/) for the current comparison matrix. The [benchmark harness](../OfficeIMO.Excel.Benchmarks/README.md) documents reproducible local runs, workload validation, allocation evidence, and data publication. Benchmark-only libraries remain isolated from the `OfficeIMO.Excel` runtime package.

## Examples

The quick start covers the smallest workbook. These examples show common read, write, reporting, and automation workflows that belong in `OfficeIMO.Excel`.

### Read rows by header

```csharp
using var reader = ExcelDocument.OpenDataReader("input.xlsx", new ExcelReadOptions {
    SheetName = "Data"
});

while (reader.Read()) {
    Console.WriteLine(reader["Name"]);
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
`LossPolicy` to `OfficeConversionLossPolicy.Allow` on conversion or save options
only after reviewing that loss. See
[XLS and XLSX compatibility](../Docs/officeimo.excel.legacy-xls-compatibility.md) for
the current capability matrix and safety contract. Use the
[migration guide](../MIGRATION.md#legacy-doc-and-xls-api-changes) for canonical API replacements.

### Import additional legacy spreadsheet formats

The `OfficeIMO.Excel` package also contains an explicit, read-only importer for selected Lotus 1-2-3, Quattro Pro, Multiplan, and Microsoft Works sources. No additional package is required, and these formats are only processed when the application calls `LegacySpreadsheetImporter` or explicitly registers the corresponding Reader handler.

```csharp
using OfficeIMO;
using OfficeIMO.Excel.Legacy;

using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import("archive.wk1");
Console.WriteLine(imported.Report.Quality);
foreach (LegacySpreadsheetCellContent cell in imported.Cells) {
    Console.WriteLine($"{cell.SheetName}!R{cell.Row}C{cell.Column}: {cell.Formula ?? cell.CachedValue}");
}
foreach (OfficeCompatibilityFinding finding in imported.Report.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

imported.Value.Save("archive.xlsx");
```

The importer never saves back to these source formats, executes macros, activates embedded objects, or resolves and refreshes external links. Each result identifies structured or salvage recovery and reports feature-level loss. Existing Excel converter packages can export the returned workbook to ODS, CSV, HTML, or PDF.

#### Profile coverage

| Family/profile | Quality | Recovered today | Explicit boundary |
| --- | --- | --- | --- |
| Lotus 1-2-3 WK1 `0x0406` record streams | Structured | empty workbooks, cells, labels, numbers, finite formula caches, bounded RPN formula translation, names, selected number formats, alignment, and chart metadata | WK1 is projected as one source sheet; other Lotus profiles remain salvage; unsupported formula tokens retain the cache with a diagnostic |
| Quattro Pro DOS WQ1/WQ2 record streams | Structured | sheet identifiers, cells, finite cached values, names, alignment, and chart metadata | Quattro formulas retain cached values with a diagnostic; WB/QPW structures, comments, advanced formatting, and live charts are not claimed |
| Microsoft Works DOS WKS `0x0404` record streams | Structured | cells, cached values, safe formulas, names, selected number formats, alignment, and chart metadata | later Works binary/compound structures and comments are not claimed |
| Later Lotus 123, Quattro QPW, and Works XLR/binary profiles | Salvage | bounded text/tabular runs and compound-content safety inventory where applicable | workbook structure, formulas, names, comments, advanced formatting, and charts are reported as unavailable |
| Microsoft Multiplan DOS 1-3 | Salvage | bounded text and tabular runs | cell zones, formulas, names, formats, comments, and charts are not yet semantically decoded |

Structured WK-derived profiles accept a valid BOF/EOF workbook with no cells and currently require ASCII text. Formula translation is allow-listed, bounded, charged against the import-wide text budget, and never evaluates the source expression. Unsupported tokens retain only a finite cached value with a loss diagnostic. `Structured` means the record stream passed the profile grammar, not that conversion is lossless; inspect `Report.Findings`, or call `imported.RequireNoLoss()` when salvage recovery, inert content, or any known approximation must fail the workflow.

### Work with XLSB workbooks

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb;

using var document = ExcelDocument.Load(
    "source.xlsb",
    new ExcelLoadOptions {
        XlsbImportOptions = new XlsbImportOptions { MaxCells = 2_000_000 }
    });

document["Data"].CellValue(2, 2, 1250m);
document.Save("edited.xlsb");

ExcelDocument.Convert("source.xlsb", "editable.xlsx");
```

XLSB detection uses package content rather than trusting the extension. The
importer projects supported values, formulas, styles, dates, geometry, views,
merges, names, and hyperlinks while retaining unknown BIFF12 records and
unmodified package parts. Supported cell edits use a native preservation-aware
rewrite. Unsupported mutations and save-time transforms fail before output is
written, so `.xlsx` bytes are never disguised as `.xlsb`.

For untrusted files, capability preflight, macro and embedded-payload handling,
and DOC/XLS/XLSB loss policies, see the
[Word and Excel interoperability guide](../Docs/officeimo.word-excel-interoperability.md).

### Stream workbook rows

List worksheet names without decoding worksheet cells, shared strings, or styles:

```csharp
IReadOnlyList<string> sheetNames = ExcelDocument.GetSheetNames(
    "input.xlsx",
    new ExcelReadOptions {
        MaxWorksheets = 256,
        MaxMetadataPartBytes = 4 * 1024 * 1024
    });
```

`GetSheetNames` supports XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, and BIFF5/BIFF8
XLS. It returns readable worksheets in workbook order and applies
`MaxInputBytes`, `MaxWorksheets`, `MaxMetadataPartBytes`, and the configured
cancellation token before worksheet data is opened.

```csharp
using OfficeIMO.Excel;

using var reader = ExcelDocument.OpenDataReader("input.xlsx", new ExcelReadOptions {
    SheetIndex = 0,
    A1Range = "A1:B1000",
    NumericAsDecimal = true
});
Console.WriteLine(reader.CurrentSheetName);
while (reader.Read()) {
    string name = reader.GetString(reader.GetOrdinal("Full Name"));
    decimal value = reader.GetDecimal(reader.GetOrdinal("Value"));
    Console.WriteLine($"{name}: {value}");
}
```

On .NET 8 and later, `ReadAsync`, `NextResultAsync`, and `RowsAsAsync<T>` propagate
cancellation through the native workbook reader:

```csharp
using OfficeIMO.Data;

await using var reader = ExcelDocument.OpenDataReader("input.xlsx", new ExcelReadOptions {
    SheetName = "Data"
});

await foreach (InvoiceRow row in reader.RowsAsAsync<InvoiceRow>(cancellationToken)) {
    await ProcessAsync(row, cancellationToken);
}
```

`ExcelDocument.OpenDataReader` returns an `ExcelWorkbookDataReader`, the package-owned read-only entry point for
XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, and BIFF8 XLS. It discovers used ranges and exposes
additional worksheets through `NextResult()`. Select one worksheet with
`SheetName` or the zero-based `SheetIndex`, and select an explicit range with
`A1Range`. `CurrentSheetName` and `CurrentSheetIndex` identify the active workbook sheet;
`CurrentResultIndex` identifies its position in the selected results. Legacy XLS is projected through
the package's existing first-party reader; use `ExcelDocument.Load` when the
workbook must be inspected, edited, or saved again. CSV provides the same typed
and ordered-parallel row-mapping contracts through the separate
`OfficeIMO.CSV` package.

On .NET 8 and later, request `DateOnly` or `TimeOnly` explicitly through
`GetFieldValue<T>` or `RowsAs<T>`. Inferred Excel date/time columns remain
`DateTime`, so moving between target frameworks does not silently change the
reader schema. Set `ExcelReadOptions.MappingErrorValuePolicy` to
`DataMappingErrorValuePolicy.Redact` when typed mapping failures must omit
source values and custom-converter exception details; the default is `Include`
for compatibility.

With `TreatDatesUsingNumberFormat` enabled, calendar formats use the workbook's
1900 or 1904 date system. Time-only and elapsed formats use an OLE Automation
`DateTime` as a carrier for the unshifted serial; changing the workbook date
system does not add days to a duration. Use a numeric getter or disable date
format interpretation when the application needs the original serial.
These rules apply to XLSX, XLS and XLSB tabular readers, including cached numeric
formulas. XLS import preserves the source date system and numeric serials rather
than reconstructing them from `DateTime` values.
Numeric properties mapped through `RowsAs<T>` or `RowsAsParallel<T>` also receive
the original serial, including cached formula values. A cell converter returning
`ExcelCellValue.NotHandled`, or a type converter returning `(false, null)`,
retains default conversion. Explicit converter results, including handled nulls,
take precedence. Type converters see the reader's original value before numeric
mapping uses the serial.
Explicit `RowMapper<T>` bindings also support aliases, optional source columns,
per-column culture, date formats and converters through
[`RowMappingColumnOptions`](../OfficeIMO.Core/README.md#per-column-row-mapping).
These controls apply to sequential, asynchronous and parallel projections.

`ExcelReadOptions.EnableWorksheetPrefetch` can overlap selected XLSX worksheet
decompression with workbook metadata parsing on a spare worker. It is disabled by default,
retains the existing package-part limits, and can be slower for small or simple workbooks.
Enable it only after measuring a representative workbook on the target CPU.

### Choose automatic or ordered parallel reads

Use `RowsAsParallel<T>()` on the public forward-only reader when typed conversion
is substantial enough to repay parallel scheduling. Workbook parsing stays
single-owner while independent row snapshots are mapped concurrently and
returned in source order:

```csharp
using OfficeIMO.Data;
using OfficeIMO.Excel;

using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
    "sales.xlsx",
    new ExcelReadOptions {
        SheetName = "Data",
        A1Range = "A1:D50001",
        InferSchema = true
    });

SalesRow[] rows = reader.RowsAsParallel<SalesRow>(
    new ParallelRowMappingOptions {
        MaxDegreeOfParallelism = 8
    }).ToArray();
```

For an `ExcelSheet`, use the explicit ordered-parallel projection API. It
supports automatic property mapping, an AOT-friendly `RowMapper<T>`, or an
`IDataRecord` factory:

```csharp
using OfficeIMO.Data;

using var document = ExcelDocument.Load("sales.xlsx");
SalesRow[] rows = document["Data"].RowsAsParallel<SalesRow>(
    "A1:D50001",
    new ParallelRowMappingOptions {
        MaxDegreeOfParallelism = 8
    }).ToArray();
```

Both public surfaces preserve source order and bound in-flight work. The
forward-only form applies to every format supported by `OpenDataReader`; the
`ExcelSheet` form projects an already loaded editable workbook and enables the
bounded schema inference required for safe snapshots. Set `InferSchema = true`
on a directly opened reader as shown above. A degree of one, or a schema with
provider-owned object/mutable fields, uses the sequential mapping contract.
Parsing and decompression are not claimed to run in parallel, and small or cheap
rows can still be faster sequentially.

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

### Plan and apply structural edits

```csharp
using var document = ExcelDocument.Load("report.xlsx");
var sheet = document["Data"];

var plan = sheet.PlanInsertColumns(firstColumn: 2, count: 2);
Console.WriteLine($"{plan.AffectedCells} cells; {plan.Impacts.Count} impact groups");
ExcelMutationResult result = plan.Apply();

sheet.InsertRowsTransactional(firstRow: 5, count: 2);
sheet.Range("A2:C20").CopyTo("E2");
sheet.Range("E2:G20").MoveTo("I2");
sheet.Range("A2:C20").TransposeTo("M2");

document.Save();
```

Transactional row, column, and cell-shift edits update workbook-owned formulas
and names, tables and filters, validations, conditional formatting, merges,
links, comments, drawings, charts, sparklines, pivot sources, print definitions,
allowed-edit ranges, and ignored-error regions. Copy, move, and transpose use the
same bounded dry-run and diagnostic contract. Existing direct row and column
methods remain available for callers that have already performed their own gate.
Formula results are marked dirty and the workbook requests recalculation on
open. Shared formulas are materialized into equivalent normal formulas before
the edit so that each member can be rewritten independently.

The operation rejects edits that would cross an array-formula boundary or
PivotTable output, remove a table header or totals row, or move a dependent
reference beyond Excel's row limit. Remove, move, or resize that owned structure
first. Configure scan, affected-cell, rollback-snapshot, and diagnostic budgets
through `ExcelMutationPlanOptions`.

### Reference-aware formulas and native cell images

`ExcelReference` parses and converts A1 and R1C1 cells, ranges, whole rows, and
whole columns and provides intersection, union, subtraction, containment, and
offset operations. `ExcelFormulaSyntaxTree` is the shared lossless rewriter used
for formulas, defined names, structured table references, chart formulas, pivot
sources, print definitions, and structural edits. `SearchFormulas` searches by
text, function, or intersecting parsed reference; formula inspection reports
authored, cached, evaluated, dirty, deferred, unsupported, and dynamic-array
state explicitly.

Use `SetInCellImage`, `GetInCellImages`, and `RemoveInCellImage` for native rich-
value images. Their metadata follows cell sorting, filtering, sizing, copying,
moving, and structural edits; they are distinct from floating drawing images.

### File-backed editing for large workbooks

```csharp
const long packageBudget = 2L * 1024 * 1024 * 1024;
using var document = ExcelDocument.OpenFileBacked(
    "large-report.xlsx",
    new ExcelLoadOptions { MaxInputBytes = packageBudget },
    cancellationToken);

document["Data"].CellValue(2, 2, "Updated");
document.Save(new ExcelSaveOptions { MaxTemporaryPackageBytes = packageBudget });
```

`OpenFileBacked` stages the editable Open XML package in an owner-only temporary
file, copies with fixed memory and deterministic cancellation, and honors load
and Open XML part limits. The normal `Load`, direct writer, streaming reader,
and unchanged-package fast paths are unchanged. XLS and XLSB projection continue
to use `Load`.

### Validation lists and typed reads

```csharp
using var document = ExcelDocument.Load("input.xlsx");
var sheet = document["Data"];

sheet.ValidationList("C2:C100", new[] { "New", "Processed", "Hold" });
sheet.Range("D2:D100").Validate.WholeNumberBetween(1, 10, errorMessage: "Use 1 through 10");

List<RowModel> rows = sheet.RowsAs<RowModel>("A1:C100").ToList();

// Omitting the range maps the populated worksheet range.
List<RowModel> populatedRows = sheet.RowsAs<RowModel>().ToList();

// For NativeAOT or explicit column control, use the same mapper shape as CSV.
List<RowModel> mappedRows = sheet.RowsAs<RowModel>(map => map
    .FromColumn<string>("Name", static (row, value) => { row.Name = value; return row; })
    .FromColumn<string>("Status", static (row, value) => { row.Status = value; return row; }))
    .ToList();

// Constructor-bound models use an explicit factory and do not require T : new().
List<ImmutableRow> immutableRows = sheet.RowsAs(factory: row => new ImmutableRow(
    row.GetString(row.GetOrdinal("Name")),
    row.GetString(row.GetOrdinal("Status"))))
    .ToList();

public sealed class RowModel {
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}

public sealed record ImmutableRow(string Name, string Status);
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

sheet.AddHistogramChart(new[] { 8d, 10d, 10d, 12d, 15d, 18d }, row: 18, column: 1, binCount: 3);
sheet.AddParetoChart(
    new[] { "Late", "Damaged", "Missing" },
    new[] { 18d, 7d, 3d },
    row: 18,
    column: 10);

document.Save();
```

Histogram, Pareto, funnel, and waterfall helpers build compatible XLSX charts from raw values. Native ChartEx authoring is available separately for funnel, waterfall, box-and-whisker, treemap, and sunburst layouts:

```csharp
var modernData = new ExcelChartData(
    new[] { "Qualified", "Proposal", "Won" },
    new[] { new ExcelChartSeries("Deals", new[] { 42d, 18d, 7d }) });

var modernChart = sheet.AddModernChart(
    modernData,
    row: 18,
    column: 10,
    chartType: ExcelModernChartType.Funnel,
    title: "Pipeline");

modernChart.SetTitle("Current pipeline")
    .SetPlacement(row: 18, column: 10, widthPixels: 640, heightPixels: 360);
```

`ExcelModernChart` can inspect imported ChartEx objects and change their name, title, supported layout, and one-cell placement without replacing unrelated markup. `UpdateData` is available only when the ChartEx formulas resolve to OfficeIMO's owned hidden chart-data sheet; visible imported business data is never claimed as writable chart storage. Other imported charts remain formatting-preserving but data replacement is rejected. Use `ExcelFormatCapabilityReport.Current.ToMarkdown()` when a workflow must choose between XLSX, XLS, and XLSB targets.

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

sheet.CellValue(5, 1, "APAC");
sheet.CellValue(5, 2, "Beta");
sheet.CellValue(5, 3, "Q2");
sheet.CellValue(5, 4, 87000);
sheet.UpdatePivotTableSource("SalesPivot", sheet, "A1:D5");
document.AddPivotSlicer(
    "SalesPivot",
    "Region",
    sheet.Name,
    new ExcelSlicerViewOptions { Name = "RegionFilter", Row = 12, Column = 8 });

var pivot = sheet.GetPivotTables().Single(p => p.Name == "SalesPivot");
Console.WriteLine($"{pivot.Name}: {string.Join(", ", pivot.RowFields)}");

var chart = sheet.ChartFromTable("SalesTable")
    .VarianceColumns("Revenue by region")
    .At(row: 12, column: 1);
chart.SetPivotSource("SalesPivot");

document.Save();
```

Pivot support covers source-range pivots, row/column/page/data fields, styles, layouts, filters, calculated fields, grouping metadata, shared-cache-aware source updates, refresh-on-open, and readback. `AddPivotSlicer` authors native slicer caches, worksheet views, and drawing anchors for supported fields. `AddPivotTimeline` does the same for date-only fields. Compatible views reuse shared caches; removing the last view can prune its cache. Unsupported imported siblings remain preserved.

Read an existing saved pivot value with `GetPivotData`. The lookup uses the pivot's
saved output cells and axis metadata; changing source cells does not refresh it.

```csharp
using var document = ExcelDocument.Load("saved-sales-pivot.xlsx");
var sheet = document.GetSheet("Summary");
ExcelCellData value = sheet.GetPivotData("SalesPivot", "Revenue",
    new Dictionary<string, object?> { ["Region"] = "North" });
```

The lightweight evaluator also supports `GETPIVOTDATA("Revenue",Summary!A1,"Region","North")`.
Both paths support saved hierarchies with multiple row and column
fields, selected page items, multiple measures, default subtotals, and text, numeric, Boolean, date, error or blank
item keys. Numeric range groups and Years/Months or Years/Quarters/Months date hierarchies on row or column axes use their displayed
group labels as criteria. A numeric criterion equal to a group's lower bound also selects that group; other raw source
numbers do not. A year criterion accepts its displayed text
or an integer year. Lookup follows saved tabular, compact and outline views, including
subtotals displayed at the top and collapsed groups. A partial selection uses its displayed subtotal, or a single matching
leaf when no subtotal is displayed. Ambiguous matches and unknown items, fields
or measures return typed errors; a blank intersection of existing items returns
zero. Lookup accepts at most 256 source fields, measures and criteria, one Values
pseudo-field across the axes, 100,000 shared keys per field, 100,001 axis/field
items (including totals or default items), and one million output cells. Each
axis also has a one-million budget for field visits and indexed criterion items.
Other grouped profiles and views without saved axis items remain
unsupported. The public method throws `NotSupportedException` for those profiles;
formula recalculation leaves them unsupported. `AddPivotTable` authors metadata
and refresh-on-open settings; it does not populate an output view for lookup.

For date items, use a `DateTime` or a numeric serial in the workbook's date system.
Use serial `60` (or its time fraction) to select Excel's fictitious February 29,
1900; a .NET `DateTime` cannot represent that date. Pivot-cache date items retain
Excel's separate OLE Automation encoding, including early-1900 values.
When a mixed field contains both a date and a number with the same serial,
lookup selects the date item, as Excel does.
To select an error item through the public API, pass
`new ExcelCellData(ExcelCellDataKind.Error, "#DIV/0!")` as the criterion value.
The string `"#DIV/0!"` selects a text item with that spelling. An error-valued
formula argument propagates its error through formula calculation.

Call `MaterializePivotTable` to generate an output view without an Office refresh:

```csharp
ExcelPivotMaterializationResult result = sheet.MaterializePivotTable("SalesPivot");
Console.WriteLine(result.OutputRange);
```

For a typed `OrderDate` column, build and populate a year/month view in the same workbook:

```csharp
sheet.Pivot("A1:B6").Rows("OrderDate").Sum("Sales", "Metric")
    .DateHierarchy("OrderDate", ExcelPivotGroupBy.Years, ExcelPivotGroupBy.Months)
    .Layout(ExcelPivotLayout.Tabular).At("D4", "SalesByMonth");
sheet.MaterializePivotTable("SalesByMonth");
```

For a text `Product` column, create a named manual group in a derived field and
materialize the resulting two-level row view:

```csharp
sheet.Pivot("A1:B6").Rows("Product").Sum("Sales", "Metric")
    .Layout(ExcelPivotLayout.Tabular).At("E4", "SalesByProductGroup");
sheet.AddPivotManualGrouping("SalesByProductGroup", "Product", "Product2",
    new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear" } },
    hiddenGroupItems: new[] { "Carrot" });
sheet.MaterializePivotTable("SalesByProductGroup");
var fruit = sheet.GetPivotData("SalesByProductGroup", "Metric",
    new Dictionary<string, object?> { ["Product2"] = "Fruit" });
```

Pass multiple entries in the dictionary to add several named groups in one derived field.
Each source item can belong to only one group; ungrouped text items remain
individual entries. The source field must be on a row or column axis. Use the
pivot's actual name, which may differ from a requested duplicate name.

Materialization supports up to 256 ordinary measures with unique captions, all eleven aggregation modes,
and multiple fields on each axis, with at most 256 cache fields in total. Numeric range groups with explicit finite
bounds and intervals, including decimal intervals, are supported when each bound and interval is at most 9 × 10¹⁵
in magnitude and the range has at most 99,998 buckets.
Years/Months and Years/Quarters/Months date hierarchies and manual text groups with saved group labels are also supported on row and column axes. Hidden items on qualified manual, numeric, and date groups are applied to source records and saved views. Grouping on page fields is not.
Date hierarchies require typed date source values. The saved group labels are retained, including labels created under
a non-English Excel locale. Renamed item captions remain visible in the generated view; `GETPIVOTDATA`
accepts both the displayed caption and the original cache key for a renamed ordinary item.
It writes a tabular view in first-seen key order for ordinary fields and group-label order for groups,
with selected page items, hidden row/column items, optional grand totals, typed values/errors, source cache
records and consistent axis metadata. It saves a copy of the source records in
the pivot cache and clears refresh-on-open. Source formulas must have saved
cached results; materialization does not calculate them.
When several pivot views share a cache, materializing any one of them regenerates
every view of that cache in one transaction. All views must qualify for
materialization and produce identical cache fields and records. The returned
`AffectedPivotTables` names the refreshed views; `OutputRange` remains the range
of the requested view.
Manual item filters retain their selected keys when source item order changes.
New source keys remain excluded from a manually filtered field unless its
`IncludeNewItemsInFilter` setting admits them. A filter that leaves no source
records is rejected before changing the saved view.
Materialization supports label equals, begins with, ends with, contains, and
their negations on ordinary row or column fields with text, Boolean, blank,
error, or numeric captions. Numeric captions use General or `#,##0` pivot-field
formatting; other numeric formats and date captions remain unqualified. It also
supports label greater/less comparisons and inclusive
between or exclusive not-between ranges when every caption and criterion uses
ASCII letters. Other text is rejected because Excel's localized ordering can
select different items. Qualified label comparisons use the current process
culture, matching Excel's locale-dependent ordering when run under the same
locale. On a single ordinary axis field, it supports value equals,
not equals, greater than, greater than or equal, less than, less than or
equal, inclusive between, exclusive not-between, and top/bottom count,
percent, and sum filters. Count filters include ties at the cutoff. Percent
and sum filters select by cumulative signed measure value, including the item
that crosses the threshold. Percent uses the current grand total; a zero
grand total retains all items. Zero and negative item aggregates are supported.
An item whose selected aggregate is an error does not enter the ranking; Count
and Count Numbers still rank their numeric results for error-valued source cells.
Ranking with no numeric aggregate is not qualified for headless materialization.
Label and value filters can share that field. Value filters use the selected measure's aggregation
over current source rows; refreshed views reevaluate item membership. The
filter definition remains available for Excel refresh, and filtered items are
not converted into manually hidden items. Filter evaluation is bounded by the
source-cell work budget.
For multiple measures, the Values field appears on exactly one axis. Values-first,
intermediate and Values-last ordering are preserved, including views without real row or column
fields. `AddPivotTable` creates this field on columns by default, or on rows when
`dataOnRows` is true. Repeated native axis prefixes are expanded during saved-view lookup.
Regeneration expands collapsed groups and normalizes compact and outline layouts to
tabular output with subtotals below their children. Hidden subtotals stay hidden;
custom subtotal functions become automatic subtotals using each measure's aggregation.
Date labels receive date/time formatting when the destination has no date format.
Date subtotal captions use `yyyy-MM-dd HH:mm:ss`, including Excel's fictitious
February 29, 1900. Existing destination date formats are retained.

The operation rejects destination collisions with unrelated values, formulas,
source cells, other pivots, tables and merged ranges. It replaces a previous
materialized view and clears its stale tail while retaining existing cell styles.
Preparation and writing run under the workbook lock; the shared mutation owner
rolls back interrupted writes or invalid generated metadata. Use
`ExcelMutationPlanOptions` to lower source/output and rollback budgets. The source
and affected output are each capped at one million cells. The measure-input budget is
source records times measure count times the number of displayed input levels on each
axis: the leaf level plus enabled intermediate subtotals. Grand totals add at most
one level per axis, bounding actual aggregate updates at four times that budget.
For shared caches, source-cell visits, measure-input visits, and affected output
cells are each summed across views against the configured budget.
Fields and measures are capped at 256; distinct shared keys per field and observed
leaf aggregate states across measures are capped at 100,000. All aggregate states,
including subtotals and grand totals, must fit the output budget. Saved axis items,
including totals, are capped at 100,001. On each axis, the sum of saved field items
and shared keys must also fit the configured budget, so all generated criteria
combinations remain within the saved-lookup indexing limit.

Other date levels and layouts, custom non-text grouping, formatted date and broader numeric label filters,
localized label range ordering, explicit show-all empty item combinations, other label/value pivot filter types and multi-field value filtering,
calculated fields, incompatible shared-cache groupings, and associated interaction
caches still require additional materialization support. Imported definitions
remain available through the existing metadata and refresh-on-open APIs.

### Guarded query-backed tables

```csharp
using var document = ExcelDocument.Create("query-report.xlsx");
var sheet = document.AddWorksheet("Results");

var query = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
    ConnectionName = "SalesQuery",
    CommandText = "sales/current",
    WorksheetName = sheet.Name,
    StartCell = "B3",
    TableName = "SalesResults",
    ColumnNames = new[] { "Region", "Amount" }
});

ExcelQueryRefreshResult refresh = await document.RefreshQueryAsync(
    query.ConnectionName,
    applicationQueryHost,
    new ExcelQueryExecutionPolicy {
        AllowExecution = true,
        MaximumRows = 100_000,
        MaximumCells = 500_000
    },
    cancellationToken);
```

OfficeIMO stores the native connection, table, and query-table relationship chain but does not ship a database or network provider. The application-owned `IExcelQueryExecutionHost` interprets the opaque command and returns rows. OfficeIMO applies row, column, cell, and character budgets before a transactional table replacement. Commands loaded from imported workbooks require the separate `AllowImportedCommands` opt-in.

### Formula inspection and calculation policy

Scalar calculation supports parentheses, unary signs, percentages, powers,
arithmetic precedence, text concatenation, comparisons, and expressions inside
supported function arguments. For example, `SUM(A1,2)*3` and
`IF(A1*2>10,"High","Low")` can be calculated without Excel. Errors such as
`#DIV/0!`, `#VALUE!`, and `#NUM!` propagate through these expressions. Syntax
and nested function evaluation are bounded to 128 levels; unsupported formulas
continue to be reported explicitly rather than calculated from stale caches.

Named references can resolve to A1 ranges or numeric, text, Boolean, and error
constants, including bounded aliases and worksheet-local scope. Arbitrary
formulas stored in defined names remain outside this calculation subset.
`ISNUMBER` distinguishes numbers from Boolean values; `ISLOGICAL` and `NA`
support typed guards and error fallbacks. `SEARCH` and text criteria recognize
`*`, `?`, and tilde escapes. `ROUND` handles decimal midpoints away from zero.
`TEXT` uses invariant English formatting and accepts an explicit `[$-409]`
locale prefix; other locale qualifiers remain unsupported.

Array calculation uses the range owned by `SetArrayFormula`, or the existing
array reference in an imported workbook. A fixed range must match the result
shape:

```csharp
using var document = ExcelDocument.Create("array-report.xlsx");
var sheet = document.AddWorksheet("Results");
sheet.SetArrayFormula("G2:H4", "SEQUENCE(3,2,10,-2)");
document.Calculate();
document.Save();
```

Use `SetDynamicArrayFormula` when the result shape may change:

```csharp
using var document = ExcelDocument.Create("dynamic-report.xlsx");
var sheet = document.AddWorksheet("Results");
sheet.CellValue(1, 1, 3);
sheet.SetDynamicArrayFormula("G2", "SEQUENCE(A1,2)");
document.Calculate(); // G2:H4 contains 1 through 6
document.Save();
```

`SEQUENCE` produces rectangular numeric results. `FILTER` accepts a row or
column mask of numbers or Booleans, including comparisons between operands of
the same native type and a scalar empty-result fallback. Blank mask cells are
false; mixed-type comparisons and text-mask coercion remain deferred.
`SORT` supports numeric sort keys, both directions, and sorting rows or columns;
blank, text, and mixed-type key collation remains deferred. `UNIQUE` preserves the first
occurrence of distinct rows or columns and supports `exactly_once`. These
functions can be nested within this array subset. Each input and output is
limited to 100,000 cells, with at most 32 array-expression levels.
`SetArrayFormula` writes the Excel compatibility prefixes for these four
functions, including nested calls, while preserving literals and reference text.

Fixed arrays update only when their result matches the authored range and the
range contains neither another formula nor a merged cell. Dynamic arrays write
Excel's native dynamic-array metadata, resize their cached range as inputs
change, and clear vacated children. A nonempty cell, fixed array, table, or
merged range that blocks a spill yields `#SPILL!` at the anchor without
overwriting the blocker. Clearing the blocker allows the next calculation to
spill. Value, formula, bulk-value, and clear operations reject edits inside an
active spill; call `ClearArrayFormula` to remove the array first. Array children
participate in same-sheet dependency calculation before caches are written;
a reference to the anchor is scalar. Unsupported expressions remain deferred
with existing caches preserved. Empty results use Excel's rich-value `#CALC!`
metadata. Zero-sized `SEQUENCE` dimensions produce `#CALC!`; negative
dimensions produce `#VALUE!`. The checked-in Excel-produced array corpus
verifies cached results through save and reopen.

Loaded dynamic arrays can resize when their saved child caches still match the
worksheet's original content. If old cells were edited outside the sheet API,
the saved source changed before calculation, the raw Open XML package was
accessed before the sheet, or the source exceeds the bounded ownership scan,
calculation reports `#SPILL!` and keeps ambiguous cells. Clear and re-author the
array only when you own the old range.

Cached reads resolve native rich-value `#CALC!` and `#SPILL!` errors across the
object model, range reads, and forward-only data readers. Unknown or unresolved
rich values retain their ordinary cell fallback. Rich error metadata is limited
to 16 MiB per part and 100,000 entries per collection; readers also honor a
smaller configured metadata limit. Repeated calculation reuses error records,
and images and other rich values keep their existing metadata and relationships.

`CellAt(...).GetValue().Value` preserves the native cached result type, including
Boolean values and text that looks numeric, such as the result of `="12"`.

Calculation honors the workbook's 1900 or 1904 date system. In the 1900 system,
`DATE` and date-part functions retain Excel's fictitious February 29, 1900.
When converting serial 60 to `DateTime`, `ExcelDateSystemConverter.FromSerial`
uses February 28 because .NET cannot represent that fictitious date. The
checked-in Excel-produced corpus verifies function composition, named and
cross-sheet references, typed errors, rounding, and both date systems through
calculation, save, and reopen.

`Calculation.MaximumDependencyDepth` defaults to 256 formula cells per chain.
The expression-nesting limit applies separately within each cell. A dependency
limit or circular reference clears affected caches and marks those formulas
dirty; unrelated formulas still calculate. Raise the dependency budget explicitly
for longer trusted chains. Calculation also checks the caller's available stack;
deeply composed dependencies can be left dirty even within those two budgets.
The independent corpus includes forward and reverse
300-cell chains and a circular-reference workbook, with cache recovery and
diagnostics verified after save and reopen. Circular caches from Excel are
preserved as input evidence rather than treated as a numeric cycle solution.

For bounded row-memory XLSX exports, set `RequireStreaming = true` in
`ExcelTabularWriteOptions`. With `WriteDataReader`, also set
`UseSharedStrings = false` and `AutoFit = false`; tables require headers.
With `WriteRows`, turn off tables and automatic sizing. Incompatible options
are rejected before source rows are consumed or destination bytes are written.

```csharp
using var document = ExcelDocument.Load("report.xlsx");

var formulas = document.InspectFormulas();
Console.WriteLine(formulas.ToMarkdown());
Console.WriteLine($"Maximum dependency depth: {formulas.DependencyGraph.MaximumDependencyDepth}");

foreach (var cycle in formulas.DependencyGraph.CircularReferences) {
    Console.WriteLine("Circular: " + string.Join(" -> ", cycle.References));
}

foreach (var formula in formulas.Formulas.Where(f => !f.IsSupportedByOfficeIMO)) {
    Console.WriteLine($"{formula.SheetName}!{formula.CellReference}: {formula.UnsupportedReason}");
}

document.Calculation.MaximumDependencyDepth = 512;
int calculated = document.Calculate();
document.Save("report.xlsx", new ExcelSaveOptions {
    EvaluateFormulasBeforeSave = true,
    ForceFullCalculationOnOpen = true
});
```

### Preflight a workbook before choosing a workflow

```csharp
using var document = ExcelDocument.Load(
    "incoming.xlsx",
    new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });

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

Use workflow preflight when an application needs to decide whether a workbook is safe for readback, cell-value edits, structure-changing edits, cached-formula reads, OfficeIMO formula calculation, template binding, or first-party PDF report export. Preserve-only features such as macros, unsupported imported interaction markup, threaded comments, external links, custom XML, OLE objects, and form controls are reported with package details instead of being silently ignored.

### Package and VBA signatures

`InspectSignatures()` and `InspectPackageSignatures(...)` remain provider-free. Pass an optional security provider only when creating a package signature or validating its cryptography, package digests, certificate chain, and revocation. OPC timestamp elements are reported as structural evidence; this shared result surface does not claim timestamp-token validation:

```csharp
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
ExcelDocument.SignPackageSignature("report.xlsx", security, signingCertificate);
OfficePackageSignatureValidationReport validation =
    ExcelDocument.ValidatePackageSignatures("report.xlsx", security);
```

`InspectVbaSignatures(...)`, `ValidateVbaSignatures(...)`, and `SignVbaProject(...)` use the managed bounded VBA core shared with Word and PowerPoint. It creates and validates legacy, agile, and V3 carriers in `.xlsm`, `.xltm`, `.xlam`, and `.xlsb` on every supported platform through an explicit `IOfficeSecurityProvider`. A registered Microsoft Office SIP is an optional Windows differential check, not a signing dependency. Sign the VBA project before applying an OPC package signature.

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

// The same lifecycle covers imported classic and Office extension rules.
var rules = sheet.GetConditionalFormattingRules("C2:D100");
var copied = sheet.CloneConditionalFormattingRule(rules[0], "E2:E100");
copied.Priority = 1;
sheet.UpdateConditionalFormattingRule(copied);

document.Save();
```

For extension-only visuals, use the same format-neutral model rather than a
second Open XML-specific API:

```csharp
sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
    Source = ExcelConditionalFormattingSource.Office2010Extension,
    Range = "F2:F100",
    Type = "DataBar",
    DataBarColor = "FF4472C4",
    DataBarBorderColor = "FF203864",
    DataBarNegativeColor = "FFC00000",
    DataBarAxisColor = "FF000000",
    DataBarBorder = true,
    DataBarGradient = false,
    DataBarThresholds = new[] {
        new ExcelConditionalFormatThreshold { Type = "AutoMin" },
        new ExcelConditionalFormatThreshold { Type = "AutoMax" }
    }
});
```

`GetConditionalFormattingRules`, `AddConditionalFormattingRule`,
`UpdateConditionalFormattingRule`, `CloneConditionalFormattingRule`,
`ReorderConditionalFormattingRules`, `RemoveConditionalFormattingRule`, and
`ClearConditionalFormatting` manage standard and Office extension rules through
one API. Common edits do not reserialize unchanged formulas or visuals, so
unrecognized imported attributes and extension children are retained. Excel
image/PDF projection emits stable diagnostics when extension semantics are
approximated or omitted; native XLS export rejects extension-only rules rather
than silently discarding them.

### Tune larger exports

```csharp
using var document = ExcelDocument.Create("large-report.xlsx");
document.Execution.Mode = ExcelExecutionMode.Automatic;
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

When rows arrive asynchronously, use the same headers and row writer without
buffering the sequence:

```csharp
await ExcelDocument.WriteRowsAsync(
    output,
    GetRowsAsync(cancellationToken),
    new[] { "Id", "Name", "Created", "Active" },
    static (writer, row) => writer
        .Write(row.Id)
        .Write(row.Name)
        .Write(row.Created)
        .Write(row.Active),
    ct: cancellationToken);
```

`WriteRowsAsync` awaits and disposes the source enumerator and writes each row
once. Because the final row count is unknown when package output starts, this
overload does not support `CreateTable` or `AutoFit`; add those features through
the editable workbook API when they are required.

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

When rows already have a fixed schema, pass a `DataTable` directly. This keeps
its column order and avoids generic object flattening. Fixed-schema tables allow
up to 5,000,000 cells by default, including the header row. Set `MaxRows`,
`MaxColumns`, or `MaxCells` in the configuration callback when a report needs a
different bounded limit. Excel's worksheet row and column limits still apply;
select or split data that exceeds them.

```csharp
using System.Data;
using OfficeIMO.Excel;

var rows = new DataTable("Members");
rows.Columns.Add("Enabled", typeof(bool));
rows.Columns.Add("AD State", typeof(string));
rows.Rows.Add(true, "Enabled");
rows.Rows.Add(false, "Disabled");

using var document = ExcelDocument.Create("members.xlsx");
document.Compose("Members", composer => {
    composer.TableFrom(
        rows,
        title: "Members",
        configure: options => options.Columns = new[] { "Enabled", "AD State" });
    composer.Finish(autoFitColumns: true);
});
document.Save();
```

## Managed image export

Ranges, worksheets, and workbook batches can be exported as PNG, JPEG, TIFF, lossless WebP, or SVG:

```csharp
using OfficeIMO.Drawing;

byte[] tiff = sheet.Range("A1:F20").ToTiff(new ExcelImageExportOptions {
    ShowGridlines = false,
    RasterEncoding = new OfficeRasterEncodingOptions {
        Tiff = new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.PackBits }
    }
});

document.ToImages()
    .ForSheets("Summary", "Data")
    .FitWithin(1600, 1200)
    .AsWebp()
    .Save("preview-images");

sheet.ToImages()
    .UsePrintArea()
    .SplitByManualPageBreaks()
    .WithGridlines(false)
    .AsPng()
    .Save("print-area-pages");
```

Excel layout, print-title, page-setup, and header/footer composition stays in `OfficeIMO.Excel`; final sizing and encoding are delegated once to `OfficeIMO.Drawing`. Worksheet batches can use print-area segments or manual page breaks without routing through a workbook export.

Charts use the same managed image contract and preserve their anchored dimensions:

```csharp
ExcelChart chart = sheet.ChartFromTable("RevenueTable")
    .RevenueTrend("Revenue trend")
    .Size(640, 320)
    .At(row: 1, column: 5);

chart.ToImage()
    .WithBackground(OfficeColor.White)
    .AsPng()
    .Save("revenue-chart.png");

chart.SaveAsSvg("revenue-chart.svg");
```

## Content provenance

`ExcelDocument.InspectProvenance("input.xlsx")` reports C2PA and AI-specific IPTC metadata in the workbook and its supported embedded images. `ExcelDocument.RemoveProvenance("input.xlsx", "clean.xlsx")` removes the selected carriers. Signed-package mutation is blocked unless `OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures` is selected explicitly. Optional cryptographic C2PA verification remains in `OfficeIMO.Security`.

When the workbook is already in memory, inspect its encoded package bytes directly. This overload validates the package and inspects supported provenance without accessing the filesystem:

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Provenance;

static OfficeProvenanceReport InspectUploadedExcel(byte[] packageBytes) =>
    ExcelDocument.InspectProvenance(packageBytes, "upload.xlsx");
```

## Concealed-content inspection and cleanup

`ExcelDocument.InspectContentSafety(...)` reports hidden sheets/rows/columns, zero geometry, tiny or transparent text, `;;;` hidden display formats, resolved low contrast, comments, hidden defined names, drawing runs/fields and ancestor groups, drawing alternative text, and Unicode evidence across XLSX/XLSM/XLSB and supported legacy XLS input. `ExcelDocument.RemoveSelectedContent(...)` clears only reviewed current payloads, isolates shared-string edits per referencing cell, and writes the original physical format. Conditional-format rendering is diagnosed rather than guessed; signed legacy XLS cleanup fails closed.

## Adjacent packages

| Package | Use it for |
| --- | --- |
| [OfficeIMO.Excel.Pdf](../OfficeIMO.Excel.Pdf/README.md) | Excel to PDF export through `OfficeIMO.Pdf`, plus PDF table import to Excel. |
| [OfficeIMO.Excel.GoogleSheets](../OfficeIMO.Excel.GoogleSheets/README.md) | Planning and exporting Excel content to Google Sheets. |
| [OfficeIMO.Excel.Benchmarks](../OfficeIMO.Excel.Benchmarks/README.md) | Benchmark harness for Excel workloads. |

## Deeper docs

- [Compatibility matrix](COMPATIBILITY.md)
- [Large workbook guidance](../Docs/officeimo.excel.large-workbook-guidance.md)
- [Repository roadmap](../Docs/ROADMAP.md)
- [Excel examples](../OfficeIMO.Examples/Excel)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** Open XML SDK for `.xlsx` package mechanics. Microsoft BCL/JSON compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Core`. The workbook API, BIFF8 `.xls` reader/writer, large-data paths, validation, and PNG/JPEG/TIFF/WebP/SVG export are first-party.
- **Security:** Open XML and VBA signature carriers are inspected and signed-package mutations fail safely without a cryptographic dependency. Package and VBA signature creation and cryptographic validation accept an explicit `IOfficeSecurityProvider`; `OfficeIMO.Security` is not pulled transitively.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 21 | 13 | 0 | 15 | 2 | 2 |
| Read | 34 | 0 | 17 | 0 | 0 | 2 |
| Edit | 11 | 0 | 39 | 2 | 1 | 0 |
| Preserve | 10 | 0 | 39 | 0 | 0 | 0 |
| Inspect | 6 | 0 | 0 | 0 | 0 | 0 |
| Validate | 4 | 0 | 0 | 0 | 0 | 2 |
| Remove | 4 | 0 | 0 | 0 | 1 | 0 |
| Convert | 48 | 39 | 0 | 9 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Excel` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
