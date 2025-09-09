# OfficeIMO.Excel — .NET Excel Utilities

OfficeIMO.Excel provides a lightweight, typed, and ergonomic API for reading and writing .xlsx files on top of Open XML. It focuses on fast values reads, editable row workflows, and write helpers that avoid extra file handles.

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

### Link a table column by header

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Read;

var s = doc["Summary"]; // table with a header row

// Find the column index of the "Domain" header
int domainCol = s.ColumnIndexByHeader("Domain");

// Build an A1 for just that column (rows 2..N)
string colLetter = A1.ColumnIndexToLetters(domainCol);
string a1 = $"{colLetter}2:{colLetter}51"; // adjust end row as needed

// Turn each cell into an internal link to a same-named sheet
s.LinkCellsToInternalSheets(a1, text => text, targetA1: "A1", styled: true);
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
- Header sugar on sheet: `sheet.SetByHeader(row, "Status", "Processed")`, `sheet.ColumnIndexByHeader("Value")`.
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

- Values-only read: available (`Rows`, `Rows("A1:C3")`, `RowsAs<T>`)
- Editable rows: available (`RowsObjects()`)
- Fluent read/write and header-wide helpers: planned

## Report Composer (generic)

`ReportSheetBuilder` provides reusable, high-level building blocks for rich worksheets: titles, sections, key–value property grids, bullet lists, tables from objects, and references. It favors performance by batching writes and using Excel tables and AutoFilter where helpful.

Example

```csharp
using var doc = ExcelDocument.Create(path);
var report = new OfficeIMO.Excel.Fluent.Report.ReportSheetBuilder(doc, "Summary");

report.Title("Scan Summary", "Generated by Tool X")
      .Section("Overview")
      .PropertiesGrid(new (string, object?)[] {
          ("Checked", 120), ("Warnings", 7), ("Errors", 1)
      }, columns: 3)
      .Section("Top Domains")
      .BulletedList(new[]{"evotec.xyz","fabrikam.net"});

// Table from objects with optional flattening/formatting
record Result(string Domain, string Area, string Status, int Score);
var data = new[]{
    new Result("example.com","Mail","Warning",7),
    new Result("evotec.xyz","Web","Ok",9),
};
report.TableFrom(data, title: "Results", configure: o => o.HeaderCase = HeaderCase.Title);

doc.Save(openExcel: true);
```

Notes
- `TableFrom<T>` uses `ObjectFlattenerOptions` so you can whitelist columns, expand nested objects, trim prefixes, and choose how collections are handled (`JoinWith` or `ExpandRows`).
- All blocks are regular Excel constructs; you can further format via `report.Sheet` using standard APIs.
- Uses the library’s `ExecutionPolicy` to scale on big datasets; override via `ExcelDocument.Execution` if needed.

## Links and Ranges

Use built-in helpers to parse A1 ranges, iterate rows/columns, and create clear, styled hyperlinks.

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Read; // A1 helpers

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

### Range variant (no table)

When you have a plain rectangular range with headers in the first row, you can link by header inside that range:

```csharp
// Headers in A1:B1 (Domain, RFC) and two data rows (A2:B3)
// Link Domain column cells to same-named sheets
s.LinkByHeaderToInternalSheetsInRange("A1:B3", "Domain", targetA1: "A1", styled: true);

// Link RFC column cells to IETF datatracker URLs (smart display when title not provided)
s.LinkByHeaderToUrlsInRange("A1:B3", "RFC", rfc => $"https://datatracker.ietf.org/doc/html/{rfc}", styled: true);
```
