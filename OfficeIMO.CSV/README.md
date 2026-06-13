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

new CsvDocument()
    .WithDelimiter(';')
    .WithHeader("Name", "Age", "City")
    .AddRow("Przemek", 36, "Mikolow")
    .AddRow("Dominika", 30, "Mikolow")
    .SortBy("Age")
    .Filter(row => row.AsString("City") == "Mikolow")
    .Save("people.csv");
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

## Boundaries

- This package owns CSV parsing, writing, transforms, and validation.
- Reader integration belongs in `OfficeIMO.Reader.Csv`.
- Excel workbook behavior belongs in `OfficeIMO.Excel`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, `net472`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
