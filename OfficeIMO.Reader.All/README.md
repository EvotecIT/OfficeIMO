# OfficeIMO.Reader.All

`OfficeIMO.Reader.All` is a thin composition package for applications that want OfficeIMO's local format handlers without registering every adapter separately.

## Install

```powershell
dotnet add package OfficeIMO.Reader.All
```

## Use the preset

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers()
    .WithMaxConcurrentReads(4)
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("input.epub");
```

The preset adds AsciiDoc, CSV/TSV, EPUB, HTML, JSON, LaTeX, OpenDocument, PDF, RTF, Visio, XML, YAML, and ZIP handlers. Word, Excel, PowerPoint, Markdown, email, and plain text remain built into `OfficeIMO.Reader`.

Configure a format through one options object:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers(new ReaderAllOptions {
        Csv = new OfficeIMO.Reader.Csv.CsvReadOptions {
            ChunkRows = 100
        },
        ZipTraversal = new OfficeIMO.Zip.ZipTraversalOptions {
            MaxEntries = 500,
            MaxTotalUncompressedBytes = 64L * 1024 * 1024
        }
    })
    .Build();
```

Registrations are copied into the builder's immutable snapshot. The preset does not mutate `DocumentReader` global state.

## Dependency boundary

This package contains no parser, provider, model, native binary, process launcher, or network client. It references OfficeIMO's existing local adapter packages and therefore carries their established managed dependency graph. OCR packages are deliberately excluded because they require an engine or executable; add one explicitly only when the host chooses it.
