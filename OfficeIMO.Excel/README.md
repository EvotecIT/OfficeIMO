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

`ExcelDocument.OpenDataReader` returns an `ExcelWorkbookDataReader`, the package-owned read-only entry point for
XLSX, XLSM, XLSB, and BIFF8 XLS. It discovers used ranges and exposes
additional worksheets through `NextResult()`. Select one worksheet with
`SheetName` or the zero-based `SheetIndex`, and select an explicit range with
`A1Range`. `CurrentSheetName` and `CurrentSheetIndex` identify the active workbook sheet;
`CurrentResultIndex` identifies its position in the selected results. Legacy XLS is projected through
the package's existing first-party reader; use `ExcelDocument.Load` when the
workbook must be inspected, edited, or saved again. CSV uses the parallel
`CsvDocument.OpenDataReader` API from the separate `OfficeIMO.CSV` package.

On .NET 8 and later, request `DateOnly` or `TimeOnly` explicitly through
`GetFieldValue<T>` or `RowsAs<T>`. Inferred Excel date/time columns remain
`DateTime`, so moving between target frameworks does not silently change the
reader schema. Set `ExcelReadOptions.MappingErrorValuePolicy` to
`DataMappingErrorValuePolicy.Redact` when typed mapping failures must omit
source values and custom-converter exception details; the default is `Include`
for compatibility.

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
    .AsWebp()
    .Save("preview-images");
```

Excel layout, print-title, page-setup, and header/footer composition stays in `OfficeIMO.Excel`; final raster encoding is delegated once to `OfficeIMO.Drawing`.

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
