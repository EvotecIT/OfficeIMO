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
- Provides forward-only `DbDataReader` access for large files without presenting a streaming document as an editable model.
- Provides ordered, bounded parallel projection for data readers and a span-backed transient-record path for decoded text on .NET 8 and later.
- Includes validated cross-library benchmark lanes with operating-system and run-mode provenance.

## Performance without giving up the document model

OfficeIMO.CSV has dedicated field-span, reusable-row, streaming `DbDataReader`,
projected-row and preformatted-text fast paths. The same package also keeps the
features expected from a document and ingestion model: schema inference and
validation, typed values, transforms, compressed files, malformed-input policy,
formula-injection protection, progress, cancellation, and diagnostics.

Runtime, CPU, input shape, quoting, encoding, storage, warm-up, and consumer behavior all matter. Use the [benchmark website](https://officeimo.com/benchmarks/) for the current hash-pinned CSV/XLSX/XLSB comparison matrix. Missing operating-system or run-mode evidence remains visible rather than being inferred. The [benchmark harness](../OfficeIMO.CSV.Benchmarks/README.md) documents the exact commands, semantic output validation, allocation evidence, and publication path.

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

For ordinary DTOs, `RowsAs<T>()` matches headers to writable properties without
requiring a range or mapping builder. Matching is case-insensitive and ignores
spaces and punctuation. OfficeIMO builds the writable-property plan once per
model type and reuses compiled assignments when the runtime supports them:

```csharp
List<Person> people = CsvDocument.Load("people.csv")
    .RowsAs<Person>()
    .ToList();
```

For a forward-only typed pipeline, project the existing `DbDataReader` surface.
This avoids building a `CsvDocument` and keeps reader lifetime explicit:

```csharp
using System.Data.Common;

using DbDataReader reader = CsvDocument.OpenDataReader("people.csv");
foreach (Person person in reader.RowsAs<Person>()) {
    Process(person);
}
```

The same automatic and explicit `RowsAs<T>` mappings work on both a materialized
`CsvDocument` and a forward-only reader. The caller owns and disposes the reader.

Use the explicit overload when assignments must be declared without reflection,
including trimming- and NativeAOT-sensitive applications. This overload still
requires `T : new()`:

```csharp
using OfficeIMO.CSV;

List<Person> people = CsvDocument.Load("people.csv")
    .RowsAs<Person>(map => map
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

For a non-positional record with a public parameterless constructor, an
assignment can return a new value from each step:

```csharp
using OfficeIMO.CSV;

var people = CsvDocument.Load("people.csv")
    .RowsAs<PersonRecord>(map => map
        .FromColumn<int>("Id", (person, value) => person with { Id = value })
        .FromColumn<string>("Name", (person, value) => person with { Name = value }))
    .ToList();

public sealed record PersonRecord {
    public int Id { get; init; }
    public string Name { get; init; } = "";
}
```

The mapper overload still requires a public parameterless constructor. For a
positional record or another constructor-bound model, use the factory overload:

```csharp
using OfficeIMO.CSV;

var people = CsvDocument.Load("people.csv")
    .RowsAs(factory: row => new PersonRecord(
        row.GetInt32(row.GetOrdinal("Id")),
        row.GetString(row.GetOrdinal("Name"))))
    .ToList();

public sealed record PersonRecord(int Id, string Name);
```

The factory receives the current `IDataRecord`; its typed getters use the CSV
reader's configured culture and schema conversions. The same overload is
available on `DbDataReader` and does not require `T : new()`.

### Ordered parallel mapping

Use `RowsAsParallel<T>()` when row conversion is substantial enough to repay
worker scheduling. One producer reads the forward-only source, workers receive
independent bounded batches, and results retain source order.
`MaxDegreeOfParallelism` bounds concurrent batches. Leave `BatchSize` unset for
the source's tuned bounded default; set it only when an application needs an
explicit throughput-versus-working-set tradeoff:

```csharp
using OfficeIMO.Data;
using System.Data.Common;

using DbDataReader reader = CsvDocument.OpenDataReader("people.csv");
Person[] people = reader.RowsAsParallel<Person>(
    new ParallelRowMappingOptions {
        MaxDegreeOfParallelism = 8
    },
    cancellationToken).ToArray();
```

The automatic, explicit `RowMapper<T>`, and `Func<IDataRecord, T>` projection
shapes all have parallel overloads. Automatic and explicit mapping snapshot
ordinary readers on the calling thread and map those bounded snapshots on
workers. Factory mapping also snapshots ordinary readers when their field
types are safe to copy, so its factory can run concurrently; readers with
provider-owned or mutable field values keep their native calling-thread
behavior. A degree of one always uses the corresponding sequential mapping
contract. Concurrent factories must not mutate unprotected shared state or
retain the transient `IDataRecord`.

On .NET 8 and later, decoded text can use the lower-overhead transient-record
API. The builder resolves headers once; its returned factory receives
span-backed fields directly from bounded parser batches:

```csharp
Person[] people = CsvDocument.ReadTextRowsAsParallel<Person>(
    csvText,
    header => {
        int id = header.GetOrdinal("Id");
        int name = header.GetOrdinal("Name");
        int age = header.GetOrdinal("Age");
        return row => new Person {
            Id = row.GetInt32(id),
            Name = row.GetString(name),
            Age = row.GetInt32(age)
        };
    },
    parallelOptions: new ParallelRowMappingOptions {
        MaxDegreeOfParallelism = 8
    },
    cancellationToken: cancellationToken).ToArray();
```

`CsvRecord`, and every span returned by it, is valid only during that factory
call. Do not retain either one. Use `CsvRecord.IsMissing(ordinal)` when a short
source row omitted a field and `CsvRecord.IsNull(ordinal)` when a present field
matches `CsvLoadOptions.NullValue`; an omitted field is deliberately not also
reported as null. The path falls back to correct sequential record access when
selected CSV options or a record shape cannot use the span-batch parser.
Parallel execution uses more working memory and thread-pool work; small or
cheap rows may be faster through the ordinary sequential API.

On .NET 8 and later, explicit `DateOnly` and `TimeOnly` targets are supported by
`RowsAs<T>`, `GetFieldValue<T>`, and `CsvColumnBuilder.AsDateOnly()` /
`AsTimeOnly()`. Default schema inference remains `DateTime`, so moving between
target frameworks does not silently change a column's inferred type.

Set `CsvLoadOptions.MappingErrorValuePolicy` to
`DataMappingErrorValuePolicy.Redact` when schema and row-mapping failures must
not include source values or custom-converter exception details. The default is
`Include` for compatibility.

## Read once or edit

Use `CsvDocument.OpenDataReader` when the caller only needs a forward-only
ADO.NET reader. CSV parsing, delimiter handling, schema inference, limits, and
stream ownership remain in `OfficeIMO.CSV`.

```csharp
using OfficeIMO.CSV;

using var reader = CsvDocument.OpenDataReader("large.csv");
while (reader.Read()) {
    int id = reader.GetInt32(reader.GetOrdinal("Id"));
    string status = reader.GetString(reader.GetOrdinal("Status"));
    Console.WriteLine($"{id}: {status}");
}
```

Use `CsvDocument` when the file must be transformed or saved again. Operations
such as `SortBy`, `Filter`, and `AddColumn` require materialized rows:

```csharp
var transformed = CsvDocument.Load("large.csv")
    .AddColumn("ImportedUtc", _ => DateTime.UtcNow)
    .Filter(row => row.AsString("Status") == "Ready")
    .SortBy(row => row.AsInt32("Id"));

transformed.Save("ready.csv");
```

The object returned by `CsvDocument.OpenDataReader` is an ADO.NET
`DbDataReader`, so it also plugs directly into
`DataTable.Load` and provider bulk-copy APIs. Enable inference when delimited
text should expose typed columns:

```csharp
using System.Data;
using System.Globalization;
using OfficeIMO.CSV;

using var reader = CsvDocument.OpenDataReader(
    "large.csv",
    new CsvLoadOptions { Culture = CultureInfo.InvariantCulture },
    new CsvDataReaderOptions {
        InferSchema = true,
        SchemaSampleSize = 1000
    });

var table = new DataTable();
table.Load(reader);
```

For a large typed import, enable bounded parallel projection on the reader.
Parsing remains single-owner, completed batches are returned in source order,
and the caller still consumes one `DbDataReader`, so the same reader can be
passed to `SqlBulkCopy` or another provider bulk-copy API:

```csharp
var schema = new CsvSchemaBuilder()
    .Column("Id").AsInt32()
    .Column("Amount").AsType(typeof(decimal))
    .Column("CreatedUtc").AsDateTime()
    .Done()
    .Build();

using var reader = CsvDocument.OpenDataReader(
    "large.csv",
    readerOptions: new CsvDataReaderOptions {
        Schema = schema,
        ParallelProcessing = new CsvDataReaderParallelOptions {
            MaxDegreeOfParallelism = 4,
            BatchSize = 4096
        }
    });
```

Parallel processing is opt-in and targets typed value conversion. String-only
readers keep their lower-overhead sequential fast path. The defaults use at
most four workers and bounded batches; tune either setting only with a
representative workload because more workers can be slower on multi-domain or
hybrid CPUs. Custom schema converters may run concurrently in this mode and
must be thread-safe; keep the reader sequential when a converter depends on
mutable single-threaded state.

`OpenDataReader` is the forward-only entry point. Use `CsvDocument.Load` when a
materialized document is required; 3.1 no longer exposes a load-mode switch.
`LoadAsync` and `SaveAsync` perform asynchronous source or destination I/O but
still materialize the document or serialized output. They are not an async CSV
cursor; `DbDataReader.Read()` remains the bounded forward-only read path.

Streaming readers also implement `ICsvDataReaderPositionMetadata`. Its
`RecordNumber` is the one-based data-record number, while
`PhysicalLineNumber` and `PhysicalEndLineNumber` identify the source lines for
the current record when the selected reader path retains that information.
Physical line values are `null` for materialized paths rather than estimated.

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

using var reader = CsvDocument.OpenDataReader("large.csv", new CsvLoadOptions {
    CancellationToken = cancellation.Token,
    ProgressReportInterval = 10_000,
    ProgressCallback = progress =>
        Console.WriteLine($"{progress.RecordsRead} records read")
});

long rowsRead = 0;
while (reader.Read()) {
    rowsRead++;
}

Console.WriteLine($"Imported {rowsRead} rows");
```

## Export options

CSV output supports null tokens, date/time formatting, UTC conversion, append, no-clobber checks, compression, quoting, encoding, and formula escaping. Formula escaping applies to text and configured text tokens; typed negative numeric values remain numeric:

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

When the source is already an `IDataReader`, write it directly to a path or
stream without introducing a second serialization path:

```csharp
using OfficeIMO.CSV;

using var reader = command.ExecuteReader();
CsvDocument.WriteDataReader("summary.csv.gz", reader, new CsvSaveOptions {
    CompressionType = CsvCompressionType.Auto
});
```

For large database exports with CPU-heavy quoting or value formatting, opt in
to ordered parallel formatting. The source reader is still consumed by one
thread; detached row batches are formatted concurrently and committed in the
original order:

```csharp
using var reader = command.ExecuteReader();
CsvDocument.WriteDataReaderParallel(
    "summary.csv.gz",
    reader,
    new CsvSaveOptions { CompressionType = CsvCompressionType.Auto },
    new CsvWriteParallelOptions {
        MaxDegreeOfParallelism = 4,
        BatchSize = 4096,
        MaximumBufferedCellsPerBatch = 1_048_576
    },
    cancellationToken);
```

The parallel writer keeps at most two batches in memory. Its effective row
count per batch is the smaller of `BatchSize` and
`MaximumBufferedCellsPerBatch / reader.FieldCount`. Every parallel write request
whose schema is wider than the cell budget is rejected before field-type
inspection, snapshot planning, or output mutation. Readers within that width
whose values require provider-owned access continue through the sequential
fallback and do not allocate parallel batches. Raise the cell budget only for a
trusted schema whose working set the application can afford; otherwise select
fewer fields or use sequential `WriteDataReader`. Custom
values and format providers used during parallel formatting must support
concurrent read-only access. For small or simply formatted exports,
`WriteDataReader` avoids the thread-pool and batch-buffering overhead and may be
faster; measure the real row shape on the target machine before choosing the
parallel path.

When the caller already has projected arrays, pass the shared schema once. The
writer validates every row width without repeating column-name validation:

```csharp
object?[][] projectedRows = {
    new object?[] { "Alpha", 10, true },
    new object?[] { "Beta", 20, false }
};

using var output = File.CreateText("summary.csv");
using var csv = new CsvRowWriter(output);
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

## Current limits and related packages

- `OfficeIMO.CSV` provides CSV parsing, writing, transforms, and validation.
- `DelimiterText` supports explicit multi-character delimiters. Delimiter auto-detection is still character-candidate based.
- Database bulk copy and provider behavior are available through DbaClientX or the consuming data-access layer rather than this package.
- Use `OfficeIMO.Reader.Csv` for unified Reader integration.
- Use `OfficeIMO.Excel` for workbook behavior.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, `net472`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** No third-party CSV engine. `System.Buffers` and .NET Framework reference assemblies support compatibility targets.
- **OfficeIMO:** `OfficeIMO.Core`. Parsing, streaming, schemas, transforms, compression, and object mapping are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
