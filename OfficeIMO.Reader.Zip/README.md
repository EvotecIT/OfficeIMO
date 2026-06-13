# OfficeIMO.Reader.Zip - ZIP reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Zip)](https://www.nuget.org/packages/OfficeIMO.Reader.Zip)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Zip?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Zip)

`OfficeIMO.Reader.Zip` registers a ZIP traversal adapter for `OfficeIMO.Reader`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Zip
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Zip;

DocumentReaderZipRegistrationExtensions.RegisterZipHandler(
    zipOptions: new ZipTraversalOptions {
        MaxEntries = 1000,
        MaxDepth = 8,
        MaxTotalUncompressedBytes = 100L * 1024L * 1024L
    });
```

## Examples

### Read supported files inside an archive

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Zip;

DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler();
DocumentReaderJsonRegistrationExtensions.RegisterJsonHandler();
DocumentReaderZipRegistrationExtensions.RegisterZipHandler(
    zipOptions: new ZipTraversalOptions {
        MaxEntries = 500,
        MaxTotalUncompressedBytes = 200L * 1024L * 1024L
    },
    readerZipOptions: new ReaderZipOptions {
        ReadNestedZipEntries = true,
        MaxNestedDepth = 2
    });

foreach (var chunk in DocumentReader.Read("evidence.zip", new ReaderOptions {
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

DocumentReaderZipRegistrationExtensions.RegisterZipHandler(
    zipOptions: new ZipTraversalOptions {
        MaxEntries = 100,
        MaxCompressionRatio = 50,
        MaxEntryUncompressedBytes = 10L * 1024L * 1024L
    },
    readerZipOptions: new ReaderZipOptions {
        ReadNestedZipEntries = false
    },
    replaceExisting: true);

var chunks = DocumentReader.Read("upload.zip").ToList();
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

- Reader adapter registration belongs here.
- ZIP traversal policy belongs in `OfficeIMO.Zip`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
