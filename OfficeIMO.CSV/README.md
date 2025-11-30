# OfficeIMO.CSV — Fluent CSV Document Model

Fluent, strongly typed, AOT-friendly CSV document model aligned with the OfficeIMO ecosystem (Word, Excel, etc.). Targets `netstandard2.0`, `net8.0`, `net9.0`, and `net472`, with streaming support for large files.

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.CSV?color=3b82f6&logo=nuget)](https://www.nuget.org/packages/OfficeIMO.CSV)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.CSV?color=22c55e&logo=nuget&label=downloads)](https://www.nuget.org/packages/OfficeIMO.CSV)
[![license](https://img.shields.io/badge/license-MIT-111827.svg)](./LICENSE)
[![targets](https://img.shields.io/badge/targets-netstandard2.0%20%7C%20net8.0%20%7C%20net9.0%20%7C%20net472-0ea5e9.svg)](#compatibility)

## Highlights
- Document-centric API: configure delimiter, header, culture, encoding once; compose transforms fluently.
- Strong typing without reflection: `Get<T>()`, helpers (`AsString/AsInt32`), explicit mapping builder for POCOs/records.
- Schema & validation: declare columns, types, required fields, custom rules; validate or throw.
- Streaming mode: lazily enumerate huge CSVs; opt-in materialization for transforms.
- AOT friendly: no `dynamic`, no codegen; trimming-safe delegates only.

### Why OfficeIMO.CSV instead of hand-rolled CSV code?
- **Predictable document model**: headers + rows are first-class, so transforms stay consistent (no ad-hoc `List<string[]>`).
- **Validation baked in**: schemas with required/optional columns, types, defaults, custom rules; catch issues before exporting/importing downstream.
- **Typed mapping without reflection**: explicit column→property assignments keep AOT and trimming happy and avoid hidden reflection costs.
- **Streaming + materialize on demand**: handle multi-GB CSVs lazily, but flip to in-memory only when you need sorting/filtering.
- **Fluent ergonomics**: chainable APIs mirror other OfficeIMO packages, reducing glue code and one-off parsers.
- **Cross-platform, legacy-friendly**: netstandard2.0 + net472 + modern TFMs in one package.

## Quick start
```csharp
using OfficeIMO.CSV;

var csv = new CsvDocument()
    .WithDelimiter(';')
    .WithHeader("Name", "Age", "City")
    .AddRow("Przemek", 36, "Mikołów")
    .AddRow("Dominika", 30, "Mikołów")
    .SortBy("Age")
    .Filter(r => r.AsString("City") == "Mikołów")
    .Save("people.csv");
```

Load & parse from text or file:
```csharp
var doc = CsvDocument.Load("input.csv", new CsvLoadOptions { Delimiter = ';' });
// or
var doc2 = CsvDocument.Parse(csvText);
```

## Transformations
- `AddRow`, `AddColumn(name, row => ...)`, `RemoveColumn(name)`
- `SortBy("Age")`, `SortBy<TKey>(r => r.Get<int>("Age"), descending: true)`
- `Filter(r => r.AsString("City") == "Mikołów")`
- `Transform(doc => ...)` for advanced scenarios

## Schema & validation
```csharp
var validated = CsvDocument.Load("input.csv")
    .EnsureSchema(schema => schema
        .Column("Id").AsInt32().Required()
        .Column("Name").AsString().Required()
        .Column("Age").AsInt32().Optional()
    )
    .ValidateOrThrow();
```
Retrieve errors without throwing:
```csharp
validated.Validate(out var errors);
```

## Typed mapping (no reflection)
```csharp
public sealed record Person(int Id, string Name, int Age, string City);

var people = CsvDocument.Load("people.csv")
    .Map<Person>(map => map
        .FromColumn<int>("Id",   (p, v) => p with { Id = v })
        .FromColumn<string>("Name", (p, v) => p with { Name = v })
        .FromColumn<int>("Age",  (p, v) => p with { Age = v })
        .FromColumn<string>("City", (p, v) => p with { City = v })
    )
    .ToList();
```

## Streaming large files
```csharp
foreach (var row in CsvDocument.Load("large.csv", new CsvLoadOptions
{
    Mode = CsvLoadMode.Stream,
    HasHeaderRow = true
}).AsEnumerable())
{
    var id = row.AsInt32("Id");
    // process lazily
}

// Need transforms? Materialize explicitly
var materialized = doc.Materialize().SortBy("Id");
```

## Options at a glance
- `CsvLoadOptions`: `Delimiter`, `HasHeaderRow` (default true), `TrimWhitespace` (default true), `AllowEmptyLines`, `Culture`, `Encoding`, `Mode` (InMemory/Stream).
- `CsvSaveOptions`: `Delimiter`, `IncludeHeader` (default true), `NewLine`, `Culture`, `Encoding`.

## Install
```bash
dotnet add package OfficeIMO.CSV
```

> Not seeing it on NuGet yet? During early development you can add a `ProjectReference` to `OfficeIMO.CSV.csproj` in this repo.

## Compatibility
- Frameworks: `netstandard2.0`, `net8.0`, `net9.0`, `net472`.
- Designed to be trimming/NativeAOT friendly on .NET 8+.
