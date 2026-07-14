# OfficeIMO.CSV - fluent CSV document model

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.CSV)](https://www.nuget.org/packages/OfficeIMO.CSV)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.CSV?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.CSV)

`OfficeIMO.CSV` is a fluent, strongly typed CSV document model aligned with the OfficeIMO ecosystem. It supports in-memory transforms, streaming reads, schemas, validation, typed mapping, and AOT-friendly explicit selectors.

## Install

```powershell
dotnet add package OfficeIMO.CSV
```

## Quick start

```csharp
using OfficeIMO.CSV;
using System.Globalization;

new CsvDocument()
    .WithDelimiter(';')
    .WithCulture(CultureInfo.InvariantCulture)
    .WithHeader("Name", "Age", "City")
    .AddRow("Przemek", 36, "Mikolow")
    .AddRow("Dominika", 30, "Mikolow")
    .AddColumn("Bucket", row => row.AsInt32("Age") >= 35 ? "Senior" : "Regular")
    .SortBy("Age")
    .Filter(row => row.AsString("City") == "Mikolow")
    .Save("people.csv", new CsvSaveOptions {
        Delimiter = ';',
        IncludeHeader = true,
        FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape,
        NewLine = "\n"
    });
```

## What it does

- Keeps headers and rows as a first-class document model instead of ad hoc string arrays.
- Loads from files, streams, or text and saves through configurable delimiter, culture, encoding, and newline options.
- Supports single-character delimiters through `Delimiter` and multi-character delimiters through `DelimiterText`.
- Reads and writes compressed CSV files with extension-based detection for gzip, deflate, Brotli, and zlib.
- Can escape formula-like values during save when producing CSV files that people will open in spreadsheet applications.
- Handles real-world import details such as duplicate headers, generated blank headers, null tokens, static metadata columns, custom date formats, comments, W3C `#Fields:` headers, and mismatched row lengths.
- Provides cancellation, progress callbacks, parse-error collection, field-length limits, quote normalization, and string interning for import pipelines.
- Supports `AddRow`, `AddColumn`, `RemoveColumn`, `SortBy`, `Filter`, and `Transform`.
- Provides schema inference and schema validation with required columns, typed columns, defaults, and custom rules.
- Maps rows to typed objects with explicit no-reflection mapping.
- Supports streaming mode for large files and explicit materialization when transforms are needed.
- Includes benchmark lanes against Dataplat/dbatools CSV, Sep, Sylvan, CsvHelper, and OfficeIMO fast paths.

## Performance without giving up the document model

OfficeIMO.CSV has dedicated field-span, reusable-row, streaming `DbDataReader`,
projected-row, and trusted-text fast paths. The same package also keeps the
features expected from a document and ingestion model: schema inference and
validation, typed values, transforms, compressed files, malformed-input policy,
formula-injection protection, progress, cancellation, and diagnostics.

The focused table compares equivalent 25,000-row wide field-span reads,
projected-array writes, `IDataReader` writes, and prepared-text writes. Benchmark
preflight parses every typed field or compares every prepared text field, so a
library cannot win by merely producing the right row shape. Lower is faster;
differences below 5% should be treated as ties rather than ranking claims.

<!-- officeimo-csv-benchmark-table:start -->
| Scenario | Variables | Host | Operation | Metric | OfficeIMO.CSV | CsvHelper | Dataplat.Dbatools.Csv | Sep | Sylvan.Data.Csv | Result |
| --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| Wide DataReader CSV write | Contract=IDataReader, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet local, Shape=wide, Snapshot=2026-07-14 | .NET 8 | Format and write rows | MeanMs | 1.00x (33ms) | n/a | 1.41x (47ms) | n/a | 0.81x (27ms) | OfficeIMO.CSV slower than Sylvan.Data.Csv |
| Wide field-span CSV read | Contract=field spans, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet local, Shape=wide, Snapshot=2026-07-14 | .NET 8 | Read every field | MeanMs | 1.00x (2ms) | n/a | n/a | 1.08x (2ms) | 4.48x (10ms) | OfficeIMO.CSV fastest |
| Wide projected-array CSV write | Contract=projected object arrays, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet local, Shape=wide, Snapshot=2026-07-14 | .NET 8 | Format and write rows | MeanMs | 1.00x (30ms) | 2.68x (82ms) | 1.48x (45ms) | n/a | n/a | OfficeIMO.CSV fastest |
| Wide validated text-row CSV write | Contract=preformatted text with escaping, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet local, Shape=wide, Snapshot=2026-07-14 | .NET 8 | Validate and write rows | MeanMs | 1.00x (17ms) | 1.35x (23ms) | 1.24x (21ms) | 1.23x (21ms) | 0.95x (16ms) | OfficeIMO.CSV tied with Sylvan.Data.Csv |
<!-- officeimo-csv-benchmark-table:end -->

These are local snapshots, not universal rankings. Runtime, CPU, input
shape, quoting, encoding, storage, warm-up, and consumer behavior all matter;
results will vary. See the [full benchmark harness](../OfficeIMO.CSV.Benchmarks/README.md)
for CsvHelper, Dataplat/dbatools, LumenWorks, Sep, Sylvan, `DataTable`, and
`DbDataReader` lanes and the exact commands used to reproduce them.

## Schema example

```csharp
var document = CsvDocument.Load("input.csv")
    .EnsureSchema(schema => schema
        .Column("Id").AsInt32().Required()
        .Column("Name").AsString().Required()
        .Column("Age").AsInt32().Optional())
    .ValidateOrThrow();
```

Collect validation errors without throwing when an import pipeline should report all bad rows:

```csharp
var document = CsvDocument.Load("input.csv")
    .EnsureSchema(schema => schema
        .Column("Id").AsInt32().Required()
        .Column("Name").AsString().Required()
        .Column("Age").AsInt32().Optional()
        .Column("Active").AsBoolean().WithDefault(true));

document.Validate(out var errors);
foreach (var error in errors) {
    Console.WriteLine($"{error.RowIndex}:{error.ColumnName} - {error.Message}");
}
```

Use `ConvertUsing` when a column needs domain-specific conversion before it becomes a `DataTable` or `IDataReader` value:

```csharp
var document = CsvDocument.Load("input.csv")
    .EnsureSchema(schema => schema
        .Column("Priority")
        .AsInt32()
        .ConvertUsing(value => string.Equals(Convert.ToString(value), "high", StringComparison.OrdinalIgnoreCase) ? 10 : 1));

DataTable table = document.ToDataTable();
```

Infer a schema from sampled rows when the incoming file should define the import contract:

```csharp
var document = CsvDocument.Load("input.csv", new CsvLoadOptions {
    DateTimeFormats = new[] { "dd-MMM-yyyy" }
});

CsvSchema inferred = document.InferSchema(sampleSize: 1000);
document.EnsureInferredSchema()
    .ValidateOrThrow();
```

## Typed mapping

Mapping is explicit and delegate-based, so it stays predictable for trimming and NativeAOT-sensitive applications.

```csharp
using OfficeIMO.CSV;

List<Person> people = CsvDocument.Load("people.csv")
    .Map<Person>(map => map
        .FromColumn<int>("Id", (person, value) => {
            person.Id = value;
            return person;
        })
        .FromColumn<string>("Name", (person, value) => {
            person.Name = value;
            return person;
        })
        .FromColumn<int>("Age", (person, value) => {
            person.Age = value;
            return person;
        })
        .FromColumn<string>("City", (person, value) => {
            person.City = value;
            return person;
        }))
    .ToList();

public sealed class Person {
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string City { get; set; } = "";
}
```

For immutable models, return a new instance from each assignment:

```csharp
using OfficeIMO.CSV;

var people = CsvDocument.Load("people.csv")
    .Map<PersonRecord>(map => map
        .FromColumn<int>("Id", (person, value) => person with { Id = value })
        .FromColumn<string>("Name", (person, value) => person with { Name = value }))
    .ToList();

public sealed record PersonRecord {
    public int Id { get; init; }
    public string Name { get; init; } = "";
}
```

## Streaming and materializing

Use streaming mode when the caller only needs forward-only row processing.

```csharp
foreach (var row in CsvDocument.Load("large.csv", new CsvLoadOptions {
    Mode = CsvLoadMode.Stream,
    HasHeaderRow = true,
    TrimWhitespace = true
}).AsEnumerable()) {
    int id = row.AsInt32("Id");
    string? status = row.AsString("Status");
    Console.WriteLine($"{id}: {status}");
}
```

Transforms such as `SortBy`, `Filter`, and `AddColumn` require in-memory mode. Materialize deliberately when that is the desired tradeoff:

```csharp
var transformed = CsvDocument.Load("large.csv", new CsvLoadOptions {
        Mode = CsvLoadMode.Stream
    })
    .Materialize()
    .AddColumn("ImportedUtc", _ => DateTime.UtcNow)
    .Filter(row => row.AsString("Status") == "Ready")
    .SortBy(row => row.AsInt32("Id"));

transformed.Save("ready.csv");
```

Use `CreateDataReader` when the next hop expects an ADO.NET reader, such as `DataTable.Load` or a provider bulk-copy API. Schema inference can expose typed columns while the rows remain forward-only:

```csharp
using System.Data;

var document = CsvDocument.Load("large.csv", new CsvLoadOptions {
    Mode = CsvLoadMode.Stream
});

using var reader = document.CreateDataReader(new CsvDataReaderOptions {
    InferSchema = true,
    SchemaSampleSize = 1000
});

var table = new DataTable();
table.Load(reader);
```

## Real-world headers

CSV exports often contain blank or repeated header names. By default, blank headers are generated as `H1`, `H2`, and duplicate names are renamed with suffixes so name-based row access stays unambiguous:

```csharp
var document = CsvDocument.Parse("Name,Name\nAlpha,Beta\n");

Console.WriteLine(string.Join(", ", document.Header));
// Name, Name_2
```

Use `DuplicateHeaderBehavior` when a pipeline needs to preserve source names exactly or reject ambiguous files:

```csharp
var strict = new CsvLoadOptions {
    DuplicateHeaderBehavior = CsvDuplicateHeaderBehavior.Throw
};

CsvDocument.Load("input.csv", strict);
```

Append static metadata columns during import when a database or audit pipeline needs source context on every row:

```csharp
var document = CsvDocument.Load("input.csv", new CsvLoadOptions {
    StaticColumns = new Dictionary<string, object?> {
        ["SourceFile"] = "input.csv",
        ["ImportedUtc"] = DateTime.UtcNow
    }
});
```

Use `NullValue` and `DateTimeFormats` when a CSV producer uses explicit null tokens or non-default date shapes:

```csharp
var document = CsvDocument.Load("input.csv", new CsvLoadOptions {
    NullValue = "<null>",
    DateTimeFormats = new[] { "dd-MMM-yyyy", "yyyyMMdd-HHmmss" }
});

DateTime created = document.AsEnumerable().First().AsDateTime("Created");
```

The parser defaults to lenient quoted-field handling for compatibility with common PowerShell CSV imports. Use strict mode when malformed quotes should fail the import:

```csharp
var document = CsvDocument.Load("input.csv", new CsvLoadOptions {
    QuoteParsingMode = CsvQuoteParsingMode.Strict
});
```

Use `DelimiterText` for multi-character delimiters such as `||` or `::`. Quoted fields can still contain the delimiter text:

```csharp
var document = CsvDocument.Parse(
    "Name||Value\nAlpha||\"one||two\"\n",
    new CsvLoadOptions { DelimiterText = "||" });

document.Save("pipes.csv", new CsvSaveOptions {
    DelimiterText = "||",
    NewLine = "\n"
});
```

Long-running import paths can opt into cancellation and progress reporting without changing the document model:

```csharp
using var cancellation = new CancellationTokenSource();

var document = CsvDocument.Load("large.csv", new CsvLoadOptions {
    Mode = CsvLoadMode.Stream,
    CancellationToken = cancellation.Token,
    ProgressReportInterval = 10_000,
    ProgressCallback = progress =>
        Console.WriteLine($"{progress.RecordsRead} records read")
});
```

## Export options

CSV output supports null tokens, date/time formatting, UTC conversion, append, no-clobber checks, compression, quoting, encoding, and formula escaping:

```csharp
CsvDocument.Load("input.csv")
    .Save("output.csv.gz", new CsvSaveOptions {
        NullValue = "<null>",
        DateTimeFormat = "yyyy-MM-ddTHH:mm:ssZ",
        UseUtc = true,
        CompressionType = CsvCompressionType.Auto,
        FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape,
        NewLine = "\n"
    });
```

Append without rewriting the header:

```csharp
CsvDocument.Load("next.csv")
    .Save("combined.csv", new CsvSaveOptions {
        Append = true,
        IncludeHeader = false
    });
```

## Objects and ad hoc data

`FromObjects` is useful for small exports from anonymous objects, DTOs, or dictionaries:

```csharp
var rows = new[] {
    new { Name = "Alpha", Count = 10, Active = true },
    new { Name = "Beta", Count = 20, Active = false }
};

CsvDocument.FromObjects(rows)
    .Save("summary.csv");
```

Use direct object writing for larger exports when the caller does not need to materialize a `CsvDocument` first. The same save options are honored, including null tokens, date/time formatting, UTC conversion, compression, append, and no-clobber checks:

```csharp
CsvDocument.SaveObjects("summary.csv.gz", rows, new CsvSaveOptions {
    NullValue = "<null>",
    DateTimeFormat = "yyyy-MM-ddTHH:mm:ssZ",
    UseUtc = true,
    CompressionType = CsvCompressionType.Auto
});
```

When the caller already has projected arrays, pass the shared schema once. The
writer validates every row width without repeating column-name validation:

```csharp
object?[][] projectedRows = {
    new object?[] { "Alpha", 10, true },
    new object?[] { "Beta", 20, false }
};

using var output = File.CreateText("summary.csv");
using var csv = new CsvObjectWriter(output);
csv.WriteRows(new[] { "Name", "Count", "Active" }, projectedRows);
```

Use `WriteTextRows` for arrays that are already culture-formatted; CSV escaping
and row-width validation still apply.

Parse text when a service receives CSV payloads without a temporary file:

```csharp
string payload = "Name,Amount\nAlpha,10\nBeta,20";

var document = CsvDocument.Parse(payload)
    .AddColumn("Currency", _ => "EUR");

string normalized = document.ToString(new CsvSaveOptions {
    Delimiter = ',',
    IncludeHeader = true
});
```

## Boundaries

- This package owns CSV parsing, writing, transforms, and validation.
- `DelimiterText` supports explicit multi-character delimiters. Delimiter auto-detection is still character-candidate based.
- Parallel CSV-to-database import is intentionally outside this package; database bulk copy and provider behavior belong in DbaClientX or the consuming data-access layer.
- Reader integration belongs in `OfficeIMO.Reader.Csv`.
- Excel workbook behavior belongs in `OfficeIMO.Excel`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, `net472`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** No third-party CSV engine. `System.Buffers` and .NET Framework reference assemblies support compatibility targets.
- **OfficeIMO:** `OfficeIMO.Drawing`. Parsing, streaming, schemas, transforms, compression, and object mapping are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
