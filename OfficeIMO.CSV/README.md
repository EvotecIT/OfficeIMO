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
        NewLine = "\n"
    });
```

## What it does

- Keeps headers and rows as a first-class document model instead of ad hoc string arrays.
- Loads from files, streams, or text and saves through configurable delimiter, culture, encoding, and newline options.
- Supports `AddRow`, `AddColumn`, `RemoveColumn`, `SortBy`, `Filter`, and `Transform`.
- Provides schema validation with required columns, typed columns, defaults, and custom rules.
- Maps rows to typed objects with explicit no-reflection mapping.
- Supports streaming mode for large files and explicit materialization when transforms are needed.

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
- Reader integration belongs in `OfficeIMO.Reader.Csv`.
- Excel workbook behavior belongs in `OfficeIMO.Excel`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, `net472`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
