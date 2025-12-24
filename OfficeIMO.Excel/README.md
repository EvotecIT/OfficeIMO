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
- Safe across tasks:
  - Multiple tasks can operate on the same `ExcelDocument`; the library coordinates writes.
  - Multiple `ExcelDocument` instances can run in parallel without interaction.

Quick setup

```csharp
using var doc = ExcelDocument.Create(path);
// Prefer all cores for compute; keep writes safe
doc.Execution.Mode = ExecutionMode.Automatic;
doc.Execution.MaxDegreeOfParallelism = Environment.ProcessorCount;
doc.Execution.OnDecision = (op, n, m) => Console.WriteLine($"[Exec] {op}: {n} → {m}");
// AutoFit with parallel compute
var s = doc.AddWorkSheet("Data");
// ... fill sheet ...
s.AutoFitColumns();
```

What to expect

- Noticeable wins on:
  - `AutoFitColumns/Rows` (thousands of rows),
  - bulk cell writes (`CellValues(...)`),
  - object→table transforms (when mapping + formatting is non‑trivial).
- Small ranges may remain sequential (overhead would dominate); thresholds are configurable.
- Exceptions are avoided in hot loops (e.g., header styling uses `TryGetColumnIndexByHeader`), so perf is stable.

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
```

### Validation lists & typed reads together
```csharp
var s = doc["Data"];
// Add a validation list for a status column
s.ValidationList("C2:C100", new[] { "New", "Processed", "Hold" });

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
- Use the validation overload to coerce or enforce rules when adding sheets:

```csharp
// Sanitize mode: fixes invalid characters, trims to 31 chars, and ensures uniqueness (adds " (2)")
var s = doc.AddWorkSheet("Q4:Revenue/Forecast?*", SheetNameValidationMode.Sanitize);
Console.WriteLine(s.Name); // => "Q4_Revenue_Forecast"

// Strict mode: throws if invalid
Assert.Throws<ArgumentException>(() => doc.AddWorkSheet("Bad:Name", SheetNameValidationMode.Strict));
```

## Colors and Styles

```csharp
using SixLabors.ImageSharp;

// Column background + bold via builder
s.ColumnStyleByHeader("Status", includeHeader: true)
 .Background(Color.Parse("#E7FFE7"))
 .Bold();

// Cell backgrounds
s.CellBackground(2, 3, Color.Parse("#FFFBE6"));
s.CellBackground(3, 3, "#FFE7E7");
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

- Values-only read: available (`Read()` fluent APIs, `Rows`, `Rows("A1:C3")`, `RowsAs<T>`)
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
  - ✅ Column/Range builders; conditional formatting (color scale, data bar)
- 🔗 Links
  - ✅ Internal/external links; smart host/title helpers; link‑by‑header (whole sheet or within range)
- 🔍 Filters & Sort
  - ✅ AutoFilter add/filter by header; conflict migration to table; multi‑column sort helpers
- 🧰 Data Quality
  - ✅ Validation lists; find/replace; header utilities (header→index, set by header)
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
