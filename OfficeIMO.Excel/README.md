# OfficeIMO.Excel — .NET Excel Utilities

OfficeIMO.Excel provides a lightweight, typed, and ergonomic API for reading and writing .xlsx files on top of Open XML. It focuses on fast values reads, editable row workflows, and write helpers that avoid extra file handles.

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Excel)](https://www.nuget.org/packages/OfficeIMO.Excel)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Excel?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Excel)

## Why OfficeIMO.Excel

- Pure .NET, cross‑platform — no COM automation, no Excel process required.
- Works directly on Open XML parts, but exposes ergonomic helpers (headers, ranges, tables, styles).
- Thread‑safe by design — scales heavy work across cores while keeping writes safe.
- Deterministic and validation‑friendly — predictable element ordering, optional Open XML validation.
- Practical guardrails — e.g., smart AutoFilter/table conflict handling; safe table naming; sensible defaults.
 - Fluent composers for rapid report building; can drop to explicit sheet APIs when needed.
 - A1 helpers and link‑by‑header utilities make “Excelish” operations straightforward.

### Thread Safety & Parallelism (How it works)

- Compute vs. apply phases:
  - Heavy work (e.g., measuring column widths, coercing values, building shared strings) runs in parallel.
  - The short “apply” phase that mutates the Open XML DOM is serialized using a document‑level lock.
- ExecutionPolicy controls behavior:
  - `doc.Execution.Mode` = `Automatic` (default), `Sequential`, or `Parallel`.
  - `Automatic` switches to parallel per operation when the workload exceeds a threshold.
  - `doc.Execution.MaxDegreeOfParallelism` caps parallelism (set to CPU count for best results).
  - Optional diagnostics callbacks: `OnDecision(op, items, mode)`, `OnTiming(op, elapsed)`.
  - `doc.Execution.SaveWorksheetAfterAutoFit = false` defers AutoFit worksheet-part saves until `Save()`/dispose, which is faster for large report exports that batch all worksheet changes.
- Safe across tasks:
  - Multiple tasks can operate on the same `ExcelDocument`; the library coordinates writes.
  - Multiple `ExcelDocument` instances can run in parallel without interaction.

Quick setup

```csharp
using var doc = ExcelDocument.Create(path);
// Prefer all cores for compute; keep writes safe
doc.Execution.Mode = ExecutionMode.Automatic;
doc.Execution.MaxDegreeOfParallelism = Environment.ProcessorCount;
doc.Execution.SaveWorksheetAfterAutoFit = false; // report-export mode: save once at the document boundary
doc.Execution.OnDecision = (op, n, m) => Console.WriteLine($"[Exec] {op}: {n} → {m}");
// AutoFit with parallel compute
var s = doc.AddWorkSheet("Data");
// ... fill sheet ...
s.AutoFitColumns();
```

Create in-memory (Stream)

```csharp
using var stream = new MemoryStream();
using (var doc = ExcelDocument.Create(stream)) {
    var sheet = doc.AddWorkSheet("Data");
    sheet.CellValue(1, 1, "Hello Stream");
}
// stream now contains the .xlsx package
stream.Position = 0;
File.WriteAllBytes("out.xlsx", stream.ToArray());
```

Create or open password-encrypted workbooks

```csharp
using var doc = ExcelDocument.Create("secure.xlsx");
var sheet = doc.AddWorkSheet("Data");
sheet.CellValue(1, 1, "Confidential");
doc.SaveEncrypted("secure.xlsx", "secret");

using var reopened = ExcelDocument.LoadEncrypted("secure.xlsx", "secret");
var value = reopened.Sheets[0].CellValue(1, 1);
```

Append to an existing table

```csharp
using var doc = ExcelDocument.Load(path);
var sheet = doc["Sales"];

var rows = new DataTable();
rows.Columns.Add("Revenue", typeof(decimal));
rows.Columns.Add("Region", typeof(string));
rows.Rows.Add(150m, "APAC");

// Columns are matched by table header by default, so source order can differ.
sheet.AppendDataTableToTable(rows, "SalesTable");
doc.Save();
```

What to expect

- Noticeable wins on:
  - simple `DataSet` exports through normal `InsertDataSet(...)` + `Save(...)`,
  - `AutoFitColumns/Rows` (thousands of rows),
  - bulk cell writes (`CellValues(...)`),
  - object→table transforms (when mapping + formatting is non‑trivial).
- Small ranges may remain sequential (overhead would dominate); thresholds are configurable.
- Exceptions are avoided in hot loops (e.g., header styling uses `TryGetColumnIndexByHeader`), so perf is stable.
- Saves automatically use fast package writers when the workbook shape is eligible. Inspect `doc.LastSaveDiagnostics` to see which writer was used or why the save fell back; use `ExcelSaveOptions.DisableFastPackageWriter` only when you explicitly want to force the standard save path.

Design choices you’ll run into

- Tables + AutoFilter: the library resolves conflicts for you (worksheet filter is migrated to the table when needed).
- Named ranges & sheet ops: sheet moves/removals re‑index local names; broken names are repaired before save.
- Deterministic ordering: element order is normalized before save to keep Excel happy and validation stable.

### AOT / Trimming notes

- Reflection-based helpers are preserved for dynamic/PowerShell usage.
- For NativeAOT/trimming, prefer explicit selectors to avoid reflection.

```csharp
// AOT-safe: explicit column selectors
var sheet = doc["Data"];
sheet.InsertObjects(people,
    ("Name",   p => p.Name),
    ("Value",  p => p.Value),
    ("Status", p => p.Status));
```

### Proof And Compatibility

- benchmark harness lives in `OfficeIMO.Excel.Benchmarks`
- large workbook guidance lives in `../Docs/officeimo.excel.large-workbook-guidance.md`
- current feature coverage matrix lives in `COMPATIBILITY.md`
- release steps live in `../Docs/officeimo.excel.release-checklist.md`

```csharp
var report = doc.InspectFeatures();

foreach (var feature in report.PreservedFeatures.Concat(report.UnsupportedFeatures)) {
    Console.WriteLine($"{feature.Name}: {feature.Count} ({feature.Note})");
}
```

## Quick Read Patterns

These helpers streamline reading Excel without extra reader boilerplate. They reuse the open `ExcelDocument` handle and infer headers/types for you.

```csharp
using OfficeIMO.Excel;

// Open workbook
using var doc = ExcelDocument.Load(path);

// Sheet access by name or 0-based index
var s1 = doc["Data"]; // case-insensitive
var s2 = doc[0];

// Values-only: iterate rows as dictionaries (UsedRange)
foreach (var row in s1.Rows()) {
    var name = (string)row["Name"];
    var val  = Convert.ToInt32(row["Value"]);
}

// Read a specific range and map to POCOs
var people = s1.RowsAs<Person>("A1:C10").ToList();

// Stream typed rows while the workbook remains open
foreach (var person in s1.RowsAsStream<Person>("A1:C100000")) {
    Console.WriteLine(person.Name);
}

// Friendly headers and explicit aliases are supported too
var summaries = s1.RowsAs<StatusSummary>("E1:G10").ToList();

// Editable rows: read → edit → save (first row = headers)
foreach (var row in s1.RowsObjects()) {
    if (row.Get<int>("Value") == 10) {
        row.Set("Status", "Processed");
    }
}
doc.Save();

public sealed class Person {
    public string Name { get; set; }
    public int    Value { get; set; }
    public string Status { get; set; }
}

public sealed class StatusSummary {
    [DisplayName("First Name")]
    public string GivenName { get; set; }

    [DataMember(Name = "Status Code")]
    public string Status { get; set; }

    [ExcelColumn("Total %", "Total Percent")]
    public int CompletionPercent { get; set; }
}
```

### Validation lists & typed reads together
```csharp
var s = doc["Data"];
// Add a validation list for a status column
s.ValidationList("C2:C100", new[] { "New", "Processed", "Hold" });

// Or use the range-level fluent API
s.Range("C2:C100").Validate.List("New", "Processed", "Hold");
s.Range("D2:D100").Validate.WholeNumberBetween(1, 10, errorMessage: "Use 1 through 10");

// Typed read back into POCOs
public sealed class RowModel { public string Name {get;set;} = ""; public string Status {get;set;} = ""; }
var rows = doc.Read().Sheet("Data").Range("A1:C10").AsObjects<RowModel>().ToList();
```

### Named ranges
```csharp
// Workbook-global
doc.SetNamedRange("GlobalArea", "'Data'!A1:B10");
// Sheet-local
var data = doc["Data"];
data.SetNamedRange("LocalStart", "A1");
// Reading a local name returns an unqualified A1 for convenience
Assert.Equal("$A$1", data.GetNamedRange("LocalStart"));
// Reading a global name returns a sheet-qualified A1
Assert.Equal("'Data'!$A$1:$B$10", doc.GetNamedRange("GlobalArea"));

// Validation modes: Sanitize (default) vs. Strict
// Name and range are both checked. Sanitize will coerce; Strict throws.
doc.SetNamedRange("123 Bad Name", "'Data'!A1:B10000000", validationMode: NameValidationMode.Sanitize); // becomes _123_Bad_Name and clamps rows
Assert.Equal("'Data'!$A$1:$B$1048576", doc.GetNamedRange("_123_Bad_Name"));
Assert.Throws<ArgumentOutOfRangeException>(() =>
    doc.SetNamedRange("BadStrict", "'Data'!A1:B10000000", validationMode: NameValidationMode.Strict));
```

### Header & Footer
```csharp
var s = doc.AddWorkSheet("Summary");
s.SetHeaderFooter(headerCenter: "Demo", headerRight: "Page &P of &N");
var logo = File.ReadAllBytes("logo.png");
s.SetHeaderImage(HeaderFooterPosition.Center, logo, widthPoints: 96, heightPoints: 32);

// Worksheet image with accessibility metadata
s.AddImage(2, 2, logo, "image/png", widthPixels: 160, heightPixels: 48,
    name: "CompanyLogo",
    altText: "Company logo")
 .LockAspectRatio()
 .SetSize(180, 54);
```

### Charts
```csharp
using OfficeIMO.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;

using var doc = ExcelDocument.Create(path);
var sheet = doc.AddWorkSheet("Summary");
doc.DefaultChartStylePreset = ExcelChartStylePreset.Default;

var data = new ExcelChartData(
    new[] { "Q1", "Q2", "Q3", "Q4" },
    new[] {
        new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }),
        new ExcelChartSeries("Target", new[] { 12d, 22d, 24d, 32d })
    });

var chart = sheet.AddChart(data, row: 1, column: 6, widthPixels: 480, heightPixels: 320,
    type: ExcelChartType.ColumnClustered, title: "Quarterly");

chart.SetLegend(C.LegendPositionValues.Right)
     .SetDataLabels(showValue: true)
     .SetSeriesFillColor(0, "4472C4")
     .SetDataLabelTextStyle(fontSizePoints: 9, color: "1F4E79")
     .SetDataLabelShapeStyle(fillColor: "FFFFFF", lineColor: "1F4E79", lineWidthPoints: 0.5);
chart.SetDataLabelLeaderLines(showLeaderLines: true, lineColor: "1F4E79", lineWidthPoints: 0.5);
chart.SetTitleTextStyle(fontSizePoints: 14, bold: true, color: "1F4E79")
     .SetLegendTextStyle(fontSizePoints: 9, color: "404040")
     .SetCategoryAxisTitle("Quarter")
     .SetValueAxisTitle("Revenue")
     .SetCategoryAxisTitleTextStyle(fontSizePoints: 10, bold: true)
     .SetValueAxisLabelTextStyle(fontSizePoints: 9, color: "404040")
     .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "C0C0C0", lineWidthPoints: 0.75)
     .SetCategoryAxisLabelRotation(45)
     .SetValueAxisTickLabelPosition(C.TickLabelPositionValues.Low);
chart.SetCategoryAxisReverseOrder()
     .SetValueAxisScale(minimum: 0, maximum: 100, majorUnit: 20, minorUnit: 10);
chart.SetValueAxisCrossing(C.CrossesValues.Maximum)
     .SetCategoryAxisCrossing(C.CrossesValues.Minimum)
     .SetValueAxisCrossBetween(C.CrossBetweenValues.Between)
     .SetValueAxisDisplayUnits(C.BuiltInUnitValues.Thousands, "Thousands USD", showLabel: true);
chart.SetChartAreaStyle(fillColor: "F2F2F2", lineColor: "404040", lineWidthPoints: 1)
     .SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "BFBFBF", lineWidthPoints: 0.75);
chart.SetSeriesTrendline(0, C.TrendlineValues.Polynomial, order: 2,
    displayEquation: true, displayRSquared: true, lineColor: "FF0000", lineWidthPoints: 1.5);

var labelTemplate = new ExcelChartDataLabelTemplate {
    ShowValue = true,
    Position = C.DataLabelPositionValues.Top,
    NumberFormat = "0.0",
    FontSizePoints = 9,
    TextColor = "404040",
    Separator = " - "
};
chart.SetSeriesDataLabelTemplate(0, labelTemplate)
     .SetSeriesDataLabelForPoint(0, 1, showValue: true, position: C.DataLabelPositionValues.OutsideEnd,
        numberFormat: "0.00")
     .SetSeriesDataLabelSeparatorForPoint(0, 1, " | ")
     .SetSeriesDataLabelTextStyleForPoint(0, 1, fontSizePoints: 11, bold: true, color: "FF0000");

// Use an existing range/table instead:
// sheet.AddChartFromRange("A1:D5", row: 8, column: 6, type: ExcelChartType.Line);
// sheet.AddChartFromTable("SalesTable", row: 8, column: 6, type: ExcelChartType.Line);
sheet.Chart("A1:D5").Line().Title("Trend").Size(480, 320).At(8, 6);
sheet.ChartFromTable("SalesTable").ColumnClustered().Title("Sales").At(8, 6);

// Dashboard recipes for common business charts:
sheet.AddRevenueTrendChart("A1:B13", row: 1, column: 6);
sheet.AddStatusBreakdownChart("D1:E6", row: 18, column: 6);
sheet.AddTopNBarChart("G1:H11", row: 35, column: 6, title: "Top Customers");
sheet.AddKpiScorecardChart("J1:K5", row: 52, column: 6, title: "KPI Scorecard");
sheet.AddContributionChart("M1:N6", row: 69, column: 6, title: "Contribution");
sheet.ChartFromTable("SalesTable").RevenueTrend("Sales Trend").At(52, 6);
sheet.ChartFromTable("VarianceBridge").VarianceWaterfall("Variance Bridge").At(86, 6);
```

```csharp
// Pivot table and pivot-source chart metadata
sheet.Pivot("A1:C100")
     .Rows("Region")
     .Columns("Product")
     .Filters("Channel")
     .Sum("Sales", "Total Sales", "$#,##0")
     .PercentOfTotal("Sales", "% of Total", "0.0%")
     .Style("PivotStyleMedium9")
     .Layout(ExcelPivotLayout.Tabular)
     .GrandTotals(rows: true, columns: true)
     .At("F2", "SalesPivot");

sheet.AddPivotTable(
    sourceRange: "A1:C100",
    destinationCell: "F2",
    name: "SalesPivot",
    rowFields: new[] { "Region" },
    dataFields: new[] {
        new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales", numberFormat: "$#,##0")
    },
    fieldOptions: new[] {
        new ExcelPivotFieldOptions("Region",
            sortType: FieldSortValues.Ascending,
            defaultSubtotal: false,
            hiddenItems: new[] { "Legacy" }),
        new ExcelPivotFieldOptions("Product",
            selectedItem: "Standard")
    },
    pageFields: new[] { "Product" },
    rowHeaderCaption: "Region",
    grandTotalCaption: "Total");

sheet.AddPivotChartFromRange("SalesPivot", "A1:C100", row: 12, column: 1,
    type: ExcelChartType.ColumnClustered, title: "Sales Pivot");
```

```csharp
// Combo chart with secondary axis
var comboData = new ExcelChartData(
    new[] { "Q1", "Q2", "Q3", "Q4" },
    new[] {
        new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }, ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary),
        new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d, 35d }, ExcelChartType.Line, ExcelChartAxisGroup.Secondary)
    });

var comboChart = sheet.AddChart(comboData, row: 10, column: 6, widthPixels: 480, heightPixels: 320,
    type: ExcelChartType.ColumnClustered, title: "Sales vs Trend");
comboChart.ApplyStylePreset()
          .SetSeriesMarker(1, C.MarkerStyleValues.Circle, size: 6, lineColor: "4472C4");
comboChart.SetValueAxisNumberFormat("0.00", sourceLinked: false, axisGroup: ExcelChartAxisGroup.Secondary)
          .SetSeriesDataLabels(1, showValue: true, position: C.DataLabelPositionValues.Top, numberFormat: "0.0");

// Scatter chart (X values come from the category column)
var scatterData = new ExcelChartData(
    new[] { "1", "2", "3", "4" },
    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d, 5d }, ExcelChartType.Scatter) });
var scatterChart = sheet.AddChart(scatterData, row: 20, column: 6, widthPixels: 480, heightPixels: 320,
    type: ExcelChartType.Scatter, title: "Scatter Sample");
scatterChart.SetScatterXAxisScale(minimum: 1, maximum: 10, majorUnit: 1, logScale: true);
scatterChart.SetScatterYAxisScale(minimum: 0, maximum: 6, majorUnit: 1);
scatterChart.SetScatterYAxisCrossing(C.CrossesValues.Minimum, crossesAt: 2d);

// Scatter chart with explicit X/Y ranges
sheet.AddScatterChartFromRanges(new[] {
    new ExcelChartSeriesRange("Series 1", "A2:A5", "B2:B5"),
    new ExcelChartSeriesRange("Series 2", "A2:A5", "C2:C5")
}, row: 30, column: 6, widthPixels: 480, heightPixels: 320,
   title: "Scatter (Ranges)");

// Bubble chart with explicit X/Y/Size ranges
sheet.AddBubbleChartFromRanges(new[] {
    new ExcelChartSeriesRange("Bubbles", "A2:A5", "B2:B5", "D2:D5")
}, row: 40, column: 6, widthPixels: 480, heightPixels: 320,
   title: "Bubble");

// Combo charts use per-series ChartType + AxisGroup to determine layout.       
```

Note: Bubble charts require explicit X/Y/size ranges via `AddBubbleChartFromRanges`. Ranges without a sheet qualifier default to the current worksheet.

### Sparklines

```csharp
sheet.Sparklines("B2:G2")
     .Column()
     .Markers()
     .HighLow()
     .Axis()
     .Color("#4472C4")
     .At("H2:H2");
```

### Link a table column by header

```csharp
using OfficeIMO.Excel; // A1 helpers are available under OfficeIMO.Excel.A1     

var s = doc["Summary"]; // table with a header row

// Find the column index of the "Domain" header (returns false when header is missing)
if (s.TryGetColumnIndexByHeader("Domain", out int domainCol))
{
    // Build an A1 for just that column (rows 2..N)
    string colLetter = A1.ColumnIndexToLetters(domainCol);
    string a1 = $"{colLetter}2:{colLetter}51"; // adjust end row as needed

    // Turn each cell into an internal link to a same-named sheet
    s.LinkCellsToInternalSheets(a1, text => text, targetA1: "A1", styled: true);
}
else
{
    // Handle the missing header scenario (log, skip, etc.)
}
```

## Fluent Read

```csharp
// Values as dictionaries
var rows = doc.Read()
              .Sheet("Data")
              .UsedRange()
              .NumericAsDecimal(true)
              .AsRows()
              .ToList();

// Map to POCOs
var people = doc.Read()
                .Sheet("Data")
                .Range("A1:C10")
                .AsObjects<Person>()
                .ToList();

// Editable rows
foreach (var row in doc.Read().Sheet("Data").UsedRange().AsEditableRows())
{
    if (row.Get<int>("Value") >= 100)
        row.Set("Status", "Hold");
    // Set a number format or formula on a specific cell
    row["Value"].NumberFormat("0.00");
}
```

## Formula Inspection And Recalculation

```csharp
var formulas = doc.InspectFormulas();
var capabilities = formulas.Capabilities;

// The lightweight evaluator supports same-sheet arithmetic plus common reporting
// functions such as SUM, AVERAGE, AVERAGEA, MINA, MAXA, COUNTIF, SUMIF, AVERAGEIF,
// COUNTIFS, SUMIFS, AVERAGEIFS, MINIFS, MAXIFS, COUNTBLANK, SUBTOTAL,
// PRODUCT, MEDIAN, LARGE, SMALL, MODE.SNGL, MODE, GEOMEAN, HARMEAN,
// AVEDEV, DEVSQ, SUMXMY2, SUMX2MY2, SUMX2PY2, SUMPRODUCT,
// bounded statistical report formulas such as SUMSQ, STDEV.S, STDEV.P, VAR.S,
// VAR.P, PERCENTILE.INC, PERCENTILE.EXC, QUARTILE.INC, QUARTILE.EXC,
// PERCENTRANK.INC, PERCENTRANK.EXC, RANK.EQ, RANK.AVG, CORREL, SLOPE,
// INTERCEPT, RSQ, and FORECAST.LINEAR,
// bounded financial report formulas such as PMT, PV, FV, NPER, and NPV,
// rounding-to-significance helpers such as MROUND, CEILING.MATH, and FLOOR.MATH,
// exact-match VLOOKUP/HLOOKUP plus MATCH/XMATCH and XLOOKUP exact/next-smaller/next-larger lookups,
// text helpers such as CONCAT, CONCATENATE, TEXT, TEXTJOIN, TEXTBEFORE, TEXTAFTER, FORMULATEXT, LEFT, RIGHT, MID,
// LEN, TRIM, SUBSTITUTE, FIND, SEARCH, VALUE, EXACT, REPT, UPPER, LOWER, and PROPER,
// including bounded TEXT number/date/time formats for report labels,
// ABS, SIGN, ROUND, ROUNDUP, ROUNDDOWN, TRUNC, INT, CEILING, FLOOR,
// including negative digit positions for report-scale rounding,
// POWER, SQRT, LN, LOG10, EXP, PI, RADIANS, DEGREES, MOD,
// ROW, COLUMN, ROWS, and COLUMNS reference-shape helpers,
// DATE, TIME, DATEVALUE, TIMEVALUE, TODAY, NOW, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, DATEDIF, YEARFRAC,
// EDATE, EOMONTH, DAYS, DAYS360, WEEKDAY, WEEKNUM, ISOWEEKNUM, NETWORKDAYS, WORKDAY, WORKDAY.INTL,
// IF/IFS/SWITCH/CHOOSE with numeric/text comparison or selector branches,
// ISBLANK/ISNUMBER/ISTEXT/ISERROR/ISERR/ISNA/ISFORMULA report guards, AND/OR/NOT comparisons, nested formulas,
// IFERROR/IFNA fallbacks returning numbers or text, and formula dependency diagnostics.
Console.WriteLine($"Formulas: {formulas.TotalFormulas}");
Console.WriteLine($"OfficeIMO-supported: {formulas.SupportedFormulas}");
Console.WriteLine($"Need Excel/cache: {formulas.UnsupportedFormulas}");
Console.WriteLine($"Dependency issues: {formulas.DependencyIssueCount}");
Console.WriteLine(capabilities.Summary);
Console.WriteLine(formulas.DependencyGraph.ToMarkdown());

foreach (var formula in formulas.Formulas.Where(f => !f.IsSupportedByOfficeIMO)) {
    Console.WriteLine($"{formula.SheetName}!{formula.CellReference}: {formula.UnsupportedReason}");
}

foreach (var formula in formulas.Formulas.Where(f => f.HasDependencyIssues)) {
    Console.WriteLine($"{formula.SheetName}!{formula.CellReference}: {string.Join("; ", formula.DependencyIssues)}");
}

Console.WriteLine(formulas.ToMarkdown());

// Unsupported formulas are preserved. Inspection reports the likely reason, such as
// unsupported functions, unsupported argument shapes, semicolon separators, text
// concatenation operators, array constants, or dependency issues.

// Calculate formulas supported by OfficeIMO's lightweight evaluator and cache results.
// Supported same-sheet formulas can depend on other supported formula cells.
// Numeric cross-sheet cell/range references such as Data!A1 and 'Data Sheet'!A1:A3
// are supported in the lightweight evaluator, as are workbook-global and sheet-local
// named ranges that point to A1 cell/range references, plus simple table structured
// references such as SalesData[Amount] and SalesData[[#Data],[Amount]]. Text
// helpers such as CONCAT, CONCATENATE, TEXT, TEXTJOIN, TEXTBEFORE, TEXTAFTER, FORMULATEXT, LEFT, RIGHT, MID, LEN,
// TRIM, SUBSTITUTE, FIND, SEARCH, VALUE, EXACT, REPT, UPPER, LOWER, and PROPER can also
// cache text results, ROW/COLUMN/ROWS/COLUMNS can cache reference-shape values,
// exact-match lookups can return text values, and IF/IFERROR/IFNA
// can cache text results for common reporting formulas.
int calculated = doc.Calculate();

// Ask Excel-compatible apps to calculate everything else on open.
doc.ConfigureFullCalculationOnOpen();

// Use guards when a workflow requires full OfficeIMO support or reliable cached reads.
doc.InspectFormulas().EnsureAllSupported();
doc.InspectFormulas().EnsureAllHaveCachedResults();
doc.Save();
```

For one-save calculation policy, use save options instead of setting persistent document defaults:

```csharp
doc.Save("report.xlsx", openExcel: false, options: new ExcelSaveOptions {
    EvaluateFormulasBeforeSave = true,
    ForceFullCalculationOnOpen = true
});
```

## Feature Inspection

```csharp
var report = doc.InspectFeatures();

Console.WriteLine($"Advanced features: {report.HasAdvancedFeatures}");

foreach (var feature in report.Features) {
    Console.WriteLine($"{feature.Category}: {feature.Name} = {feature.Count} [{feature.SupportLevel}]");
    foreach (var detail in feature.Details) {
        Console.WriteLine($"  - {detail}");
    }
}

Console.WriteLine(report.ToMarkdown());
report.EnsureNoUnsupportedFeatures();

// Fail fast for workflow-specific risk. For example, a data refresh job may
// reject macro-enabled templates or external workbook links before saving.
report.EnsureNoFeatures("VBA macros", "External workbook links");
report.EnsureNoFeatures(ExcelFeatureSupportLevel.Preserved);
```

## CSV And JSON Exchange

```csharp
var data = doc["Data"];

// Range/table export
DataTable table = data.ToDataTable("A1:C100");
string csv = data.ToCsv("A1:C100");
string json = data.ToJson("A1:C100");

// Import directly into a worksheet
data.FromCsv("Name,Amount\nAlpha,10\nBeta,20");
data.FromJson("[{\"Name\":\"Alpha\",\"Amount\":10},{\"Name\":\"Beta\",\"Amount\":20}]",
    startRow: 5,
    startColumn: 1);
```

## Template Markers

```csharp
// Existing workbook cells can contain text such as "Invoice {{Invoice.Number}}"
// or "Customer: {{Customer.Name}}".
int replacements = doc.ApplyTemplate(new Dictionary<string, object?> {
    ["Invoice.Number"] = "INV-001",
    ["Customer.Name"] = "Adatum",
    ["Total"] = 123.45
});

// Public properties can be used directly; nested objects become dotted markers.
doc.ApplyTemplate(invoiceModel);

// Format aliases and custom .NET formats are supported in markers:
// {{Total:currency}}, {{Completion:percent}}, {{Issued:yyyy-MM-dd}},
// {{Duration:duration}}
var options = new ExcelTemplateOptions { ThrowOnMissing = true }
    .AddFormatter("upper", (value, provider) =>
        Convert.ToString(value, provider as CultureInfo ?? CultureInfo.CurrentCulture)?.ToUpperInvariant() ?? string.Empty)
    .AddFormatter("fallback", (value, provider) =>
        value == null ? "n/a" : Convert.ToString(value, provider as CultureInfo ?? CultureInfo.CurrentCulture) ?? string.Empty);

// Custom formatters can be used as {{Customer.Name:upper}} or {{Notes:fallback}}.
doc.ApplyTemplate(invoiceModel, options);

// If a marker owns the whole cell, supported aliases write typed values and apply
// Excel number formats instead of replacing with display text.
// A cell containing only {{Total:currency}} becomes a numeric currency cell.

// Inspect marker requirements before binding:
var template = doc.InspectTemplate(invoiceModel);
foreach (var missing in template.MissingMarkerNames) {
    Console.WriteLine($"Missing template value: {missing}");
}

foreach (var marker in template.Markers) {
    Console.WriteLine($"{marker.SheetName}!{marker.CellReference}: {marker.Name} -> {marker.BoundValueKind}");
}

template.EnsureAllMarkersBound();
Console.WriteLine(template.ToMarkdown());

// Use throwOnMissing when templates must be fully bound.
doc.ApplyTemplate(values, throwOnMissing: true);

// Optional fields can be cleared instead of leaving the marker text in place.
doc.ApplyTemplate(values, new ExcelTemplateOptions {
    MissingValueBehavior = ExcelTemplateMissingValueBehavior.EmptyString
});

// A single worksheet row can be used as a repeating template row. Additional
// rows are inserted below the template row and each item is bound to one row.
sheet.ApplyTemplateRows(12, lineItems, new ExcelTemplateOptions {
    FormatProvider = CultureInfo.GetCultureInfo("en-US"),
    MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
});

// A worksheet can be repeated into one generated sheet per model. The first
// generated sheet reuses the template worksheet; additional sheets copy the
// worksheet structure, including table definitions, external hyperlinks,
// static images, generated charts with style/package/drawing parts, and legacy
// comments, and then bind markers.
doc.ApplyTemplateSheets(
    "Region Template",
    regions,
    (region, index) => region.Name,
    new ExcelTemplateOptions {
        FormatProvider = CultureInfo.GetCultureInfo("en-US"),
        MissingValueBehavior = ExcelTemplateMissingValueBehavior.Throw
    });

// Optional row sections can be kept and bound, or removed while following rows
// shift up.
sheet.ApplyTemplateOptionalRows(20, rowCount: 2, include: invoice.HasDiscount, invoice, options);
if (!invoice.HasNotes) {
    sheet.RemoveTemplateOptionalRows(24, rowCount: 3);
}

// Whole-cell markers can bind images into the drawing layer.
sheet.CellAt(2, 6).SetValue("{{Logo}}");
doc.ApplyTemplate(new Dictionary<string, object?> {
    ["Logo"] = ExcelTemplateImage.FromBytes(
        logoBytes,
        contentType: "image/png",
        widthPixels: 96,
        heightPixels: 32,
        name: "Logo",
        altText: "Company logo")
});

// Streams and URLs are supported too:
// ExcelTemplateImage.FromStream(stream, "image/png", widthPixels: 96, heightPixels: 32)
// ExcelTemplateImage.FromUrl("https://example.com/logo.png", widthPixels: 96, heightPixels: 32)
```

## Comments And Notes

```csharp
sheet.SetComment("A1", "Review total", author: "Alice", initials: "AA");
sheet.SetComment("B2", "Review status", author: "Alice", initials: "AA");

var reviewNotes = sheet.FindComments(new ExcelCommentFilter {
    Author = "Alice (AA)",
    TextContains = "Review",
    A1Range = "A1:B10"
});

sheet.UpdateComments(new ExcelCommentFilter {
    TextContains = "status"
}, "Status reviewed", author: "Carol", initials: "CC");

sheet.ClearComments(new ExcelCommentFilter {
    Author = "Alice (AA)",
    A1Range = "A1:A1"
});
```

## Data Operations

```csharp
var s = doc["Data"];

// AutoFilter: add and filter by header value
s.AutoFilterAdd("A1:C100");
s.AutoFilterByHeaderEquals("Status", new[] { "Processed", "Hold" });

// Contains filter (text):
s.AutoFilterByHeaderContains("Name", "Al");

// Sort (values-only rewrite)
s.SortUsedRangeByHeader("Value", ascending: false);
s.SortUsedRangeByHeaders(("Value", false), ("Name", true));

// Validation list
s.ValidationList("C2:C100", new[] { "New", "Processed", "Hold" });

// Find/Replace
var first = s.FindFirst("Beta");
int changed = s.ReplaceAll("New", "Processed");
```

## Saving & Validation

- Element order and sheet dimensions are normalized on save to avoid Excel repair prompts.
- Optional save options let you enable defined-name repairs and package validation.

```csharp
// Safe repair + OpenXML validation (throws on validation errors)
doc.Save("report.xlsx", openExcel: false, options: new ExcelSaveOptions {
    SafeRepairDefinedNames = true,
    ValidateOpenXml = true,
    SafePreflight = true // removes empty containers, drops orphaned refs
});

// Async variant
await doc.SaveAsync("report.xlsx", false, new ExcelSaveOptions {
    SafeRepairDefinedNames = true,
    ValidateOpenXml = true
}, ct);
```

## Sheet Names

- Excel allows up to 31 characters; disallows : \ / ? * [ ] and duplicate names (case-insensitive).
- `AddWorkSheet(...)` sanitizes by default; use the validation overload when you want to opt into `Strict` errors or explicit `None` behavior:

```csharp
var auto = doc.AddWorkSheet("Q4:Revenue/Forecast?*");
Console.WriteLine(auto.Name); // => "Q4_Revenue_Forecast"

// Sanitize mode: fixes invalid characters, trims to 31 chars, and ensures uniqueness (adds " (2)")
var s = doc.AddWorkSheet("Q4:Revenue/Forecast?*", SheetNameValidationMode.Sanitize);
Console.WriteLine(s.Name); // => "Q4_Revenue_Forecast"

// Strict mode: throws if invalid
Assert.Throws<ArgumentException>(() => doc.AddWorkSheet("Bad:Name", SheetNameValidationMode.Strict));
```

## Colors and Styles

```csharp
using OfficeIMO.Drawing;

// Column background + bold via builder
s.ColumnStyleByHeader("Status", includeHeader: true)
 .Background(OfficeColor.Parse("#E7FFE7"))
 .Bold();

// Cell backgrounds
s.CellBackground(2, 3, OfficeColor.Parse("#FFFBE6"));
s.CellBackground(3, 3, "#FFE7E7");

// Cell/range presets for common report formatting
s.CellAt(1, 1).SetValue("Amount").HeaderStyle();
s.CellAt(2, 1).SetValue(123.45).Currency(culture: CultureInfo.GetCultureInfo("en-US")).Success();
s.Range("B2:B20").Percent(decimals: 1).Warning();
s.Range("C2:C20").Date().MutedText();
```

Notes:
- `Rows()` materializes dictionaries using the first row of the range as headers.
- `RowsObjects()` returns editable row handles; setting `cell.Value` or calling `row.Set(header, value)` writes to the sheet.
- All helpers share a single open file handle; no extra opens.
- Header sugar on sheet: `sheet.SetByHeader(row, "Status", "Processed")`, `sheet.TryGetColumnIndexByHeader("Value", out var columnIndex)`.
- Prefer decimals? Use `ExcelReadPresets.DecimalFirst()` or set `new ExcelReadOptions { NumericAsDecimal = true }`.

### Column formatting by header

Use the discoverable builder to apply formats by header:

```csharp
// using System.Globalization;
var s = doc["Data"];

// Numbers
s.ColumnStyleByHeader("Value").Number(decimals: 2);
s.ColumnStyleByHeader("Count").Integer();

// Percent & currency
s.ColumnStyleByHeader("Rate").Percent(decimals: 1);
s.ColumnStyleByHeader("Amount").Currency(decimals: 2, culture: CultureInfo.GetCultureInfo("en-US"));

// Dates & durations
s.ColumnStyleByHeader("When").Date("yyyy-mm-dd");
s.ColumnStyleByHeader("When").DateTime("yyyy-mm-dd hh:mm:ss");
s.ColumnStyleByHeader("Duration").DurationHours();

// Custom Excel number format
s.ColumnStyleByHeader("Misc").NumberFormat("0.00E+00");
```

## Status

- Values-only read: available (`Read()` fluent APIs, `Rows`, `Rows("A1:C3")`, `RowsAs<T>`, `RowsAsStream<T>`)
- Editable rows: available (`RowsObjects()` / `Read().AsEditableRows()`)
- Fluent write: available (`Compose(...)`, `AsFluent().Sheet(...)`)

## Fluent Compose (write)

Two options are available for building worksheets fluently:

- Concise composer via `doc.Compose(...)` + `SheetComposer`
- Advanced builder via `doc.AsFluent().Sheet(...)` + `SheetBuilder`

### 1) Concise composer

```csharp
using OfficeIMO.Excel;

using var doc = ExcelDocument.Create(path);

doc.Compose("Report", c =>
{
    c.Title("Demo Report", "Subtitle");
    c.Callout("info", "Heads up", "Generated via fluent API");
    c.Section("Summary");
    c.PropertiesGrid(new (string, object?)[] {
        ("Author", "Tester"),
        ("Date", DateTime.Today.ToString("yyyy-MM-dd"))
    });

    var items = new[] {
        new { Name = "Alice", Score = 90, Status = "OK" },
        new { Name = "Bob",   Score = 80, Status = "Warning" }
    };

    c.TableFrom(items, title: "Scores", visuals: v => {
        v.NumericColumnDecimals["Score"] = 0;
        v.TextBackgrounds["Status"] = new System.Collections.Generic.Dictionary<string,string>
        { { "Warning", "#FFF3CD" } };
    });

    c.References(new[] { "https://example.com" });
    c.HeaderFooter(h => h.Center("Demo Report").FooterRight("Page &P of &N"));
    c.Finish(autoFitColumns: true);
});

doc.Save();
```

What it does
- Writes a title/subtitle, a callout, a key–value properties grid
- Builds a table from objects with per‑column visuals
- Adds a References section and a simple header/footer

### 2) Advanced builder

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

using var doc = ExcelDocument.Create(path);

doc.AsFluent()
   .Sheet("Data", s =>
   {
       // From objects → header + data rows
       s.RowsFrom(items, o => {
           o.HeaderCase = HeaderCase.Title;
           o.CollectionMode = CollectionMode.JoinWith;
       });

       // Turn it into a table and style it
       s.Table(t => {
           t.Name("Results");
           t.Style(TableStyle.TableStyleMedium2);
           t.IncludeAutoFilter(true);
       });

       // Column styles by header
       s.Columns(cols =>
       {
           cols.ByHeader("Score").Integer();
           cols.ByHeader("Status").Background("#FFF3CD");
       });
   })
   .End();

doc.Save();
```

### Selecting and ordering POCO columns

When converting objects to rows, you can control what columns appear and in which order:

- `Include(params string[])` – keep only these properties (full path or last segment).
- `Exclude(params string[])` – drop these properties.
- `PinFirst(params string[])` – pin specific columns to the front in the given order.
- `PinLast(params string[])` – push specific columns to the end in the given order.
- `PriorityOrder(params string[])` – set relative order after pinned‑first (1..N by position).

All methods are chainable. These are convenience wrappers over `IncludeProperties`, `ExcludeProperties`, `PinnedFirst`, `PinnedLast`, and `PropertyPriority`.

```csharp
s.RowsFrom(users, o => {
    o.ExpandProperties.Add(nameof(User.Address));
    o.HeaderCase = HeaderCase.Title;
    o.Include(nameof(User.Id), nameof(User.FirstName), nameof(User.LastName), "Address.City")
     .Exclude(nameof(User.Email))
     .PinFirst(nameof(User.Id))
     .PriorityOrder(nameof(User.LastName), nameof(User.FirstName), "Address.City")
     .PinLast(nameof(User.Email));
});

// Same for SheetComposer
composer.TableFrom(users, title: "Users", configure: o =>
{
    // Either chain them...
    o.PinFirst(nameof(User.Id))
     .PriorityOrder(nameof(User.LastName), nameof(User.FirstName), "Address.City")
     .PinLast(nameof(User.Email));
    // ...or use a single call
    // o.Order(
    //     pinFirst: new[] { nameof(User.Id) },
    //     priority: new[] { nameof(User.LastName), nameof(User.FirstName), "Address.City" },
    //     pinLast: new[] { nameof(User.Email) }
    // );
    o.HeaderCase = HeaderCase.Title;
    o.ExpandProperties.Add(nameof(User.Address));
});
```

## Links and Ranges

Use built-in helpers to parse A1 ranges, iterate rows/columns, and create clear, styled hyperlinks.

```csharp
using OfficeIMO.Excel; // A1 helpers are available via OfficeIMO.Excel.A1

var s = doc["Summary"]; // sheet

// Parse A1 → bounds
var (r1,c1,r2,c2) = A1.ParseRange("B2:D10"); // 2,2,10,4
var bounds = s.GetRangeBounds("A2:A51");

// Iterate rows/columns
s.ForEachRow("A2:A10", r => s.CellBold(r, 1, true));
s.ForEachColumn("A1:E1", c => s.CellBold(1, c, true));

// Internal links: turn a column of names into links to same-named sheets
s.LinkCellsToInternalSheets("A2:A51", text => text, targetA1: "A1", styled: true);

// Even simpler: link by header name (auto-detect rows)
s.LinkByHeaderToInternalSheets("Domain", targetA1: "A1", styled: true);

// External links with smart display (Title → RFC #### → host)
s.SetHyperlinkSmart(5, 1, "https://datatracker.ietf.org/doc/html/rfc7208"); // displays "RFC 7208"
s.SetHyperlinkHost(6, 1, "https://learn.microsoft.com/office/open-xml/");     // displays host only
s.SetHyperlink(7, 1, "https://example.org", display: "Example", style: true);

// Internal link to Summary top
s.SetInternalLink(2, 1, "'Summary'!A1", display: "Summary", style: true);
```

## Print Setup & TOC

```csharp
// Print area and titles
var sheet = doc["Data"];
doc.SetPrintArea(sheet, "A1:H100");
doc.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null);

// Table of Contents sheet with named ranges included
doc.AddTableOfContents(sheetName: "TOC", placeFirst: true, withHyperlinks: true, includeNamedRanges: true);
```

## Link Helpers by Header

```csharp
var summary = doc["Summary"];
// Link a column by header name to same-named internal sheets
summary.LinkByHeaderToInternalSheets("Domain", targetA1: "A1", styled: true);
// Or within a specific range
summary.LinkByHeaderToInternalSheetsInRange("A1:B20", "Domain", targetA1: "A1", styled: true);
// URL mapping by header
summary.LinkByHeaderToUrlsInRange("A1:B20", "RFC", rfc => $"https://datatracker.ietf.org/doc/html/{rfc}", styled: true);
```

## Conditional Formatting

```csharp
var data = doc["Data"];
// 2‑color scale over a range
data.AddConditionalColorScale("C2:C100", "#FFF0F0", "#70AD47");
// Data bar
data.AddConditionalDataBar("D2:D100", "#5B9BD5");

// Range-level fluent API
data.Range("C2:C100").ConditionalFormat.ColorScale("#FFF0F0", "#70AD47");
data.Range("D2:D100").ConditionalFormat.DataBar("#5B9BD5");
data.Range("E2:E100").ConditionalFormat.Top(10);
```

## Feature Matrix

- 📘 Workbook & Core
  - ✅ Create/Load/Save (sync/async); deterministic save ordering; optional validation
  - ✅ ExecutionPolicy (Automatic/Sequential/Parallel) with diagnostics hooks
- 📥 Reading
  - ✅ Used range detection; A1 range reads; typed reads (`RowsAs<T>()`); editable rows (`RowsObjects()`); range enumeration
- ✍️ Writing
  - ✅ Cells & ranges; object→table (`RowsFrom<T>()`); Excel table builder with AutoFilter
  - ✅ Named ranges (global & sheet‑local); header/footer text + images; print area/titles; freeze panes
  - ✅ TOC generator and back links
- 🎨 Styles & Formatting
  - ✅ Number formats (integer/decimal/percent/currency/date/datetime/custom); alignment; background fills
  - ✅ Column/Range builders; conditional formatting (color scale, data bar, icon set, top/bottom, duplicate/formula rules); range-level fluent formatting builders
- 🔗 Links
  - ✅ Internal/external links; smart host/title helpers; link‑by‑header (whole sheet or within range)
- 🔍 Filters & Sort
  - ✅ AutoFilter add/filter by header; conflict migration to table; multi‑column sort helpers
- 🧰 Data Quality
  - ✅ Validation lists and numeric/date/time/text/custom validation; range-level fluent validation builders; find/replace; header utilities (header→index, set by header)
- 🚀 Performance
  - ✅ AutoFit Columns/Rows and bulk writes leverage multi‑core compute phase

## Namespaces (updated)

- A1 helpers are under `OfficeIMO.Excel.A1` (no `OfficeIMO.Excel.Read` import needed for A1).
- Fluent read entrypoint: `doc.Read()` → `OfficeIMO.Excel.Fluent.ExcelFluentReadWorkbook`.
- Fluent write entrypoints:
  - `doc.Compose(...)` uses `OfficeIMO.Excel.Fluent.SheetComposer` internally.
  - `doc.AsFluent()` returns `OfficeIMO.Excel.Fluent.ExcelFluentWorkbook` (advanced builder APIs).

<!-- (No migration notes: these APIs are new additions.) -->

### Examples: Two Styles (Excelish vs Classic)

The repo ships two parallel Excel report examples so you can choose the style that fits your project:

- Excelish blocks (SheetComposer + helpers) — fast to author; consistent visuals; fewer foot‑guns.
  - File: `OfficeIMO.Examples/Excel/DomainDetective.Report.Sheets.cs`
  - Uses helpers such as `SectionLegend`, `KpiRow`, `Columns(...)`, `PrintDefaults`, `LinkByHeaderToInternalSheets*`.

- Classic explicit build — minimal sugar; shows standard calls step‑by‑step.
  - File: `OfficeIMO.Examples/Excel/DomainDetective.Report.Sheets.Classic.cs`
  - Does legends, bullets, and print options with direct cell writes and sheet methods.

Key helper snippets

```csharp
// Legend block (Status | Meaning | Action)
composer.SectionLegend(
    title: "Legend",
    headers: new[] { "Status", "Meaning", "Recommended Action" },
    rows: new[] {
        new[] { "OK", "Acceptable", "None" },
        new[] { "Warning", "Needs attention", "Review" },
        new[] { "Error", "Blocking", "Fix" },
    },
    firstColumnFillByValue: StatusPalettes.Default.FillHexMap);

// Side‑by‑side layout (3 columns)
composer.Columns(3, cols => {
    cols[0].Section("Totals").KeyValues(new[]{ ("Items", 120), ("Errors", 2) });
    cols[1].Section("Status").KeyValues(new[]{ ("OK", 100), ("Warning", 18), ("Error", 2) });
    cols[2].Section("Tips").BulletedList(new[]{ "Filter headers", "Click links" });
});

// Print polish (gridlines off, fit to width, landscape, narrow margins, repeat header row)
composer.PrintDefaults(fitToWidth: 1)
        .Orientation(ExcelPageOrientation.Landscape)
        .Margins(ExcelMarginPreset.Narrow)
        .RepeatHeaderRows(1, 1);
```

Tip: Prefer the Excelish style for velocity and consistency. Use the Classic example when you want fine‑grained control or to show exact underlying calls.

### Logos and Images (Headers/Footers and In‑Sheet)

- Header/Footer logos via builder:

```csharp
var logoPath = Path.Combine(AppContext.BaseDirectory, "Assets", "OfficeIMO.png");
byte[] logo = File.ReadAllBytes(logoPath);

composer.HeaderFooter(h =>
{
    h.Left("Report Title").Right("Page &P of &N");
    h.RightImage(logo, widthPoints: 96, heightPoints: 32); // header logo
    // h.FooterCenterImage(logo); // footer instead
});
```

- In‑sheet logo anchored to a cell (first page):

```csharp
composer.ImageFromUrlAt(row: 1, column: 6, url: "https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png", widthPixels: 120, heightPixels: 40);
```

To show a logo only on the second page, place a manual page break and anchor the image near the top row after the break, or keep it in the header and show a different first page (DifferentFirstPage) without the &G picture placeholder.

### Range variant (no table)

When you have a plain rectangular range with headers in the first row, you can link by header inside that range:

```csharp
// Headers in A1:B1 (Domain, RFC) and two data rows (A2:B3)
// Link Domain column cells to same-named sheets
s.LinkByHeaderToInternalSheetsInRange("A1:B3", "Domain", targetA1: "A1", styled: true);

// Link RFC column cells to IETF datatracker URLs (smart display when title not provided)
s.LinkByHeaderToUrlsInRange("A1:B3", "RFC", rfc => $"https://datatracker.ietf.org/doc/html/{rfc}", styled: true);
```
