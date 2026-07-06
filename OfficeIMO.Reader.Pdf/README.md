# OfficeIMO.Reader.Pdf - PDF reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Pdf)](https://www.nuget.org/packages/OfficeIMO.Reader.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Pdf)

`OfficeIMO.Reader.Pdf` registers a PDF adapter for `OfficeIMO.Reader` using the `OfficeIMO.Pdf` logical read model.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Pdf
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();

IReadOnlyList<ReaderChunk> chunks = DocumentReader
    .Read("invoice.pdf")
    .ToList();
```

## Examples

### Read page-aware Markdown chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();

foreach (var chunk in DocumentReader.Read("manual.pdf")) {
    Console.WriteLine($"Page {chunk.Location.Page}: {chunk.Id}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read a stream with input limits

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();

using var stream = File.OpenRead("large-report.pdf");
var chunks = DocumentReader.Read(stream, "large-report.pdf", new ReaderOptions {
    MaxChars = 12_000,
    MaxInputBytes = 100L * 1024L * 1024L
}).ToList();

foreach (var chunk in chunks.Where(chunk => chunk.Diagnostics != null)) {
    Console.WriteLine($"{chunk.Id}: {chunk.Diagnostics!.TableCount} table(s)");
}
```

### Register alongside other ingestion adapters

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Zip;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();
DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();
DocumentReaderZipRegistrationExtensions.RegisterZipHandler();

var chunks = DocumentReader.ReadFolder("KnowledgeBase", new ReaderFolderOptions {
    Recurse = true,
    DeterministicOrder = true,
    MaxFiles = 500
}).ToList();
```

## What it emits

- Page-aware chunks with `ReaderLocation.Page`.
- Markdown text, logical tables, column profiles, table diagnostics, and confidence signals.
- Source/security/form/catalog-metadata/open-action/active-content counters in `ReaderChunk.Diagnostics`.
- Document metadata for XMP, output intents, tagged structure, optional content/layers, attachments, security/signatures, navigation, links, forms, annotations, and passive actions.
- Passive action summaries without executable payloads.
- Image placeholders and visual geometry when available.
- Link annotations and typed form fields when available.

## Boundaries

- Reader adapter registration belongs here.
- PDF parsing, logical readback, and safety behavior belongs in `OfficeIMO.Pdf`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
