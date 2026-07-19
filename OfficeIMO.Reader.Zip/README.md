# OfficeIMO.Reader.Zip - ZIP reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Zip)](https://www.nuget.org/packages/OfficeIMO.Reader.Zip)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Zip?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Zip)

`OfficeIMO.Reader.Zip` provides a ZIP traversal adapter for `OfficeIMO.Reader.Core`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Zip
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Zip;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddZipHandler(new ZipTraversalOptions {
        MaxEntries = 1000,
        MaxDepth = 8,
        MaxTotalUncompressedBytes = 100L * 1024L * 1024L
    })
    .Build();
```

## Examples

### Read supported files inside an archive

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Zip;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddCsvHandler()
    .AddJsonHandler()
    .AddZipHandler(
        zipOptions: new ZipTraversalOptions {
        MaxEntries = 500,
        MaxTotalUncompressedBytes = 200L * 1024L * 1024L
        },
        readerZipOptions: new ReaderZipOptions {
        ReadNestedZipEntries = true,
        MaxNestedDepth = 2
        })
    .Build();

foreach (var chunk in reader.Read("evidence.zip", new ReaderOptions {
    MaxInputBytes = 200L * 1024L * 1024L
})) {
    Console.WriteLine($"{chunk.Kind}: {chunk.Location.Path}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Surface archive warnings

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Zip;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddZipHandler(
        zipOptions: new ZipTraversalOptions {
        MaxEntries = 100,
        MaxCompressionRatio = 50,
        MaxEntryUncompressedBytes = 10L * 1024L * 1024L
        },
        readerZipOptions: new ReaderZipOptions {
        ReadNestedZipEntries = false
        })
    .Build();

var chunks = reader.Read("upload.zip").ToList();
foreach (string warning in chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>())) {
    Console.WriteLine(warning);
}
```

## What it emits

- Safe entry enumeration through `OfficeIMO.Zip`.
- Best-effort entry extraction into `ReaderChunk`.
- Warning chunks for skipped or failed entries.
- Bounded nested ZIP traversal with `ReaderZipOptions`.
- Path and stream dispatch, including non-seekable stream support.

## Boundaries

- Reader adapter configuration belongs here.
- ZIP traversal policy belongs in `OfficeIMO.Zip`.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** None beyond platform compression APIs.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.Zip` own safe traversal, nested-archive limits, extraction projection, and warnings.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
