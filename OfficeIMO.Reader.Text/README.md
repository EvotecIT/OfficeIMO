# OfficeIMO.Reader.Text - structured text compatibility adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Text)](https://www.nuget.org/packages/OfficeIMO.Reader.Text)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Text?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Text)

`OfficeIMO.Reader.Text` is a compatibility orchestrator for structured text adapters.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Text
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Text;

DocumentReaderTextRegistrationExtensions.RegisterStructuredTextHandler(replaceExisting: true);
```

## Examples

### Keep one compatibility registration for CSV, JSON, and XML

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Text;

DocumentReaderTextRegistrationExtensions.RegisterStructuredTextHandler(
    new StructuredTextReadOptions {
        CsvChunkRows = 100,
        JsonChunkRows = 100,
        XmlChunkRows = 100,
        JsonMaxDepth = 16
    },
    replaceExisting: true);

foreach (string path in Directory.EnumerateFiles("imports")) {
    foreach (var chunk in DocumentReader.Read(path)) {
        Console.WriteLine($"{chunk.Kind}: {chunk.Id}");
    }
}
```

## What it delegates

- `.csv` and `.tsv` to `OfficeIMO.Reader.Csv`.
- `.json` to `OfficeIMO.Reader.Json`.
- `.xml` to `OfficeIMO.Reader.Xml`.

## Boundaries

- New integrations should prefer the dedicated CSV, JSON, and XML adapters.
- This package keeps a legacy registration entry point for existing consumers.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
