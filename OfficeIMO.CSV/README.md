# OfficeIMO.CSV

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.CSV)](https://www.nuget.org/packages/OfficeIMO.CSV)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.CSV?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.CSV)

Read, write and transform CSV in .NET with a typed document model. `OfficeIMO.CSV` handles in-memory edits, streaming reads, schemas and validation, typed mapping and compressed files, with AOT-friendly explicit selectors.

## Install

```sh
dotnet add package OfficeIMO.CSV
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/CSV/Program.cs" region="QuickStart" -->
```csharp
using System.Globalization;
using OfficeIMO.CSV;

CsvDocument.Create()
    .WithDelimiter(';')
    .WithCulture(CultureInfo.InvariantCulture)
    .WithHeader("Name", "Age", "City")
    .AddRow("Adele", 36, "Krakow")
    .AddRow("Lee", 30, "Krakow")
    .AddColumn("Bucket", row => row.AsInt32("Age") >= 35 ? "Senior" : "Regular")
    .SortBy("Age")
    .Save("People.csv", new CsvSaveOptions {
        Delimiter = ';',
        FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
    });
```
<!-- /powerforge:code -->

## What it does

- **A document model** for headers and rows, loaded from files, streams or text, with configurable delimiters, culture, encoding and newlines.
- **Transform** with `AddRow`, `AddColumn`, `RemoveColumn`, `SortBy`, `Filter` and `Transform`.
- **Schemas**: infer or declare required and typed columns, defaults and custom rules, and validate.
- **Typed mapping** to objects without reflection.
- **Stream large files** forward-only through a `DbDataReader`, with bounded parallel projection.
- **Real-world imports**: duplicate or blank headers, null tokens, comments, W3C `#Fields:` headers, mismatched rows, gzip, deflate, Brotli and zlib.
- **Safe exports**: escape formula-like values for files people open in spreadsheets.
- **Tab-separated values** with the same API.

The [CSV guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.csv-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| Excel, both directions | [OfficeIMO.Excel.Csv](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.Csv/README.md) |

## Limits

Multi-character delimiters use `DelimiterText`; a streaming reader isn't an editable document. The [CSV guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.csv-guide.md#current-limits-and-related-packages) lists the current limits.

## Learn more

- [OfficeIMO.CSV on officeimo.com](https://officeimo.com/products/csv/) and the [CSV docs](https://officeimo.com/docs/csv/).
- [CSV guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.csv-guide.md).
- Reading for search and AI: [OfficeIMO.Reader.Csv](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Reader.Csv/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core`, plus `System.Memory` on .NET Standard 2.0 and .NET Framework.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |
| Validate | 1 | 0 | 0 | 0 | 0 | 0 |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.CSV` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
