# OfficeIMO.Reader - document extraction facade

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader)](https://www.nuget.org/packages/OfficeIMO.Reader)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader)

`OfficeIMO.Reader` is a read-only facade for deterministic document extraction. It normalizes supported source files into `ReaderChunk` objects for search, indexing, chat, RAG, migration, and review workflows.

## Install

```powershell
dotnet add package OfficeIMO.Reader
```

## Quick start

```csharp
using OfficeIMO.Reader;

foreach (var chunk in DocumentReader.Read(@"C:\Docs\Policy.docx")) {
    Console.WriteLine(chunk.Id);
    Console.WriteLine(chunk.Location.HeadingPath);
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

## Streams and folders

```csharp
using OfficeIMO.Reader;

using var stream = File.OpenRead(@"C:\Docs\Policy.docx");
var chunksFromStream = DocumentReader.Read(stream, "Policy.docx").ToList();

var folderChunks = DocumentReader.ReadFolder(
    folderPath: @"C:\Docs",
    folderOptions: new ReaderFolderOptions {
        Recurse = true,
        MaxFiles = 500,
        MaxTotalBytes = 500L * 1024 * 1024,
        SkipReparsePoints = true,
        DeterministicOrder = true
    },
    options: new ReaderOptions {
        MaxChars = 8_000
    }).ToList();
```

## What it reads

Built-in and modular adapters can extract:

- Word (`.docx`, `.docm`) as Markdown chunks.
- Excel (`.xlsx`, `.xlsm`) as table chunks and optional Markdown previews.
- PowerPoint (`.pptx`, `.pptm`) as slide-aligned chunks, optionally including notes.
- Markdown (`.md`, `.markdown`) as parser-aware heading chunks.
- PDF, Visio, HTML, CSV/TSV, JSON, XML, EPUB, ZIP, and structured text through modular adapter packages.

## Modular adapters

Install and register only the adapters you need:

```csharp
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Zip;

DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler();
DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler();
DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();
DocumentReaderJsonRegistrationExtensions.RegisterJsonHandler();
DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();
DocumentReaderVisioRegistrationExtensions.RegisterVisioHandler();
DocumentReaderXmlRegistrationExtensions.RegisterXmlHandler();
DocumentReaderZipRegistrationExtensions.RegisterZipHandler();
```

## Host examples

### Capability discovery

```csharp
using OfficeIMO.Reader;

var capabilities = DocumentReader.GetCapabilities();
foreach (var capability in capabilities) {
    Console.WriteLine($"{capability.Id}: {string.Join(", ", capability.Extensions)}");
}

string manifestJson = DocumentReader.GetCapabilityManifestJson();
```

### Register a custom handler

```csharp
using OfficeIMO.Reader;

DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
    Id = "custom-log",
    DisplayName = "Custom log reader",
    Kind = ReaderInputKind.Text,
    Extensions = new[] { ".log" },
    ReadPath = (path, options, cancellationToken) => {
        string text = File.ReadAllText(path);
        return new[] {
            new ReaderChunk {
                Id = "log:1",
                Kind = ReaderInputKind.Text,
                Text = text,
                Location = new ReaderLocation { Path = path }
            }
        };
    }
});
```

## Host contracts

- `ReaderOptions` controls chunk size, table row limits, footnotes/notes, Excel ranges, Markdown heading chunking, hashes, and input budgets.
- `ReaderFolderOptions` controls recursion, file limits, byte limits, reparse-point handling, and deterministic folder order.
- `DocumentReader.GetCapabilities()` and `GetCapabilityManifestJson()` expose a stable host-discovery surface.
- Custom handlers can be registered with `DocumentReader.RegisterHandler(...)`.

## Boundaries

- `OfficeIMO.Reader` owns the shared extraction contract and built-in facade.
- Source-specific parsing belongs in the source package or modular adapter.
- Adapters should use `ReaderInputLimits` so input size and stream behavior stays consistent.
- AI or database storage belongs in the consuming application.

## Related packages

- [OfficeIMO.Reader.Pdf](../OfficeIMO.Reader.Pdf/README.md)
- [OfficeIMO.Reader.Visio](../OfficeIMO.Reader.Visio/README.md)
- [OfficeIMO.Reader.Html](../OfficeIMO.Reader.Html/README.md)
- [OfficeIMO.Reader.Csv](../OfficeIMO.Reader.Csv/README.md)
- [OfficeIMO.Reader.Json](../OfficeIMO.Reader.Json/README.md)
- [OfficeIMO.Reader.Xml](../OfficeIMO.Reader.Xml/README.md)
- [OfficeIMO.Reader.Epub](../OfficeIMO.Reader.Epub/README.md)
- [OfficeIMO.Reader.Zip](../OfficeIMO.Reader.Zip/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
