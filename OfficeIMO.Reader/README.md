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

- Word (`.docx`, `.docm`, `.doc`) as Markdown chunks.
- Excel (`.xlsx`, `.xlsm`, `.xls`) as table chunks and optional Markdown previews.
- PowerPoint (`.pptx`, `.pptm`) as slide-aligned chunks, optionally including notes.
- Markdown (`.md`, `.markdown`) as parser-aware heading chunks.
- PDF, RTF, Visio, HTML, CSV/TSV, JSON, XML, YAML, EPUB, ZIP, and structured text through modular adapter packages.

## Modular adapters

For services and concurrent hosts, build an isolated reader with only the adapters you need:

```csharp
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Yaml;
using OfficeIMO.Reader.Zip;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddCsvHandler()
    .AddEpubHandler()
    .AddHtmlHandler()
    .AddJsonHandler()
    .AddPdfHandler()
    .AddRtfHandler()
    .AddVisioHandler()
    .AddXmlHandler()
    .AddYamlHandler()
    .AddZipHandler()
    .Build();

var chunks = reader.Read(@"C:\Docs\data.json").ToList();
```

`Build()` freezes the handler configuration. The resulting `OfficeDocumentReader` is safe to reuse across concurrent reads, is unaffected by later builder changes, and cannot see handlers registered on another reader. The static `DocumentReader.RegisterHandler(...)` and adapter `Register...Handler()` methods remain available for compatibility, but they update a process-wide registry.

## Host examples

### Capability discovery

```csharp
using OfficeIMO.Reader;

var reader = new OfficeDocumentReaderBuilder().Build();
var capabilities = reader.GetCapabilities();
foreach (var capability in capabilities) {
    Console.WriteLine($"{capability.Id}: {string.Join(", ", capability.Extensions)}");
}

string manifestJson = reader.GetCapabilityManifestJson();
```

### Register a custom handler

```csharp
using OfficeIMO.Reader;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddHandler(new ReaderHandlerRegistration {
    Id = "custom-audit",
    DisplayName = "Custom audit reader",
    Kind = ReaderInputKind.Text,
    Extensions = new[] { ".auditx" },
    ReadPath = (path, options, cancellationToken) => {
        string text = File.ReadAllText(path);
        return new[] {
            new ReaderChunk {
                Id = "audit:1",
                Kind = ReaderInputKind.Text,
                Text = text,
                Location = new ReaderLocation { Path = path }
            }
        };
    }
})
    .Build();
```

Handlers that already expose a structured document model can register rich result delegates instead of rebuilding that model as chunks first:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddHandler(new ReaderHandlerRegistration {
    Id = "custom-rich-reader",
    DisplayName = "Custom rich reader",
    Kind = ReaderInputKind.Text,
    Extensions = new[] { ".rich" },
    ReadDocumentPath = (path, options, cancellationToken) => ReadRichDocument(path),
    ReadDocumentStream = (stream, sourceName, options, cancellationToken) => ReadRichDocument(stream, sourceName)
})
    .Build();
```

`reader.ReadDocument(...)` dispatches directly to these delegates. Existing `reader.Read(...)` calls remain usable by projecting the returned result's `Chunks` collection. A handler may continue to register `ReadPath` and `ReadStream` when chunk production is its native contract.

## Host contracts

- `ReaderOptions` controls chunk size, table row limits, footnotes/notes, Excel ranges, Markdown heading chunking, hashes, and input budgets.
- `ReaderFolderOptions` controls recursion, file limits, byte limits, reparse-point handling, and deterministic folder order.
- `OfficeDocumentReader.GetCapabilities()` and `GetCapabilityManifestJson()` expose the frozen configuration of that reader instance.
- Capability records distinguish basic path/stream support from native rich-result support through `SupportsDocumentPath` and `SupportsDocumentStream`.
- `OfficeDocumentReaderBuilder.AddHandler(...)` is the recommended custom-handler path for services and concurrent hosts.
- Static `DocumentReader` registration is retained as a process-wide compatibility surface.

## Boundaries

- `OfficeIMO.Reader` owns the shared extraction contract and built-in facade.
- Source-specific parsing belongs in the source package or modular adapter.
- Adapters should use `ReaderInputLimits` so input size and stream behavior stays consistent.
- AI or database storage belongs in the consuming application.

## Related packages

- [OfficeIMO.Reader.Pdf](../OfficeIMO.Reader.Pdf/README.md)
- [OfficeIMO.Reader.Rtf](../OfficeIMO.Reader.Rtf/README.md)
- [OfficeIMO.Reader.Visio](../OfficeIMO.Reader.Visio/README.md)
- [OfficeIMO.Reader.Html](../OfficeIMO.Reader.Html/README.md)
- [OfficeIMO.Reader.Csv](../OfficeIMO.Reader.Csv/README.md)
- [OfficeIMO.Reader.Json](../OfficeIMO.Reader.Json/README.md)
- [OfficeIMO.Reader.Xml](../OfficeIMO.Reader.Xml/README.md)
- [OfficeIMO.Reader.Yaml](../OfficeIMO.Reader.Yaml/README.md)
- [OfficeIMO.Reader.Epub](../OfficeIMO.Reader.Epub/README.md)
- [OfficeIMO.Reader.Zip](../OfficeIMO.Reader.Zip/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
