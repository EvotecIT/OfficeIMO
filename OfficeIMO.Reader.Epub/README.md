# OfficeIMO.Reader.Epub - EPUB reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Epub)](https://www.nuget.org/packages/OfficeIMO.Reader.Epub)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Epub?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Epub)

`OfficeIMO.Reader.Epub` bridges `OfficeIMO.Epub` output into `OfficeIMO.Reader.Core` chunk contracts.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Epub
```

## Configure

```csharp
using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEpubHandler(new EpubReadOptions {
        PreferSpineOrder = true,
        IncludeRawHtml = false,
        MaxChapters = 100
    })
    .Build();
```

## Examples

### Read chapters as Reader chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEpubHandler()
    .Build();

foreach (var chunk in reader.Read("book.epub", new ReaderOptions {
    MaxChars = 6_000,
    ComputeHashes = true
})) {
    Console.WriteLine($"{chunk.Id}: {chunk.Location.Path}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read from a stream and surface warnings

```csharp
using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEpubHandler(new EpubReadOptions {
        FallbackToHtmlScan = true,
        MaxChapterBytes = 2L * 1024L * 1024L
    })
    .Build();

await using var stream = File.OpenRead("upload.epub");
var chunks = reader.Read(stream, "upload.epub").ToList();

foreach (string warning in chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>())) {
    Console.WriteLine(warning);
}
```

### Read chapters and packaged resources as one rich result

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEpubHandler(new EpubReadOptions {
        MaxResources = 500,
        MaxResourceBytes = 4L * 1024L * 1024L,
        MaxTotalResourceBytes = 32L * 1024L * 1024L
    })
    .Build();
OfficeDocumentReadResult document = reader.ReadDocument("book.epub");

foreach (OfficeDocumentPage chapter in document.Pages) {
    Console.WriteLine($"{chapter.Number}. {chapter.Name}");
}

foreach (OfficeDocumentAsset asset in document.Assets) {
    Console.WriteLine($"{asset.Kind}: {asset.FileName} ({asset.LengthBytes} bytes)");
}
```

The rich reader requests bounded chapter HTML and manifest payloads from `OfficeIMO.Epub` so it can reuse the HTML semantic mapping. Images, audio, video, fonts, stylesheets, scripts, media overlays, and other manifest resources are projected as assets; content documents and navigation files are not duplicated as assets. Remote resources retain metadata but are never fetched. Audio and video are exposed for downstream processing, not played or transcribed.

Manifest assets keep their package identity and payload while inheriting useful metadata from chapter placements, such as accessible image names, titles, and dimensions. The first available placement supplies missing metadata; each visual still retains its own placement content and source path.

## What it emits

- Chapter-to-chunk projection.
- Max-character chunk splitting.
- Markdown and text chunk payloads.
- Warning chunks propagated from EPUB parser warnings.
- Virtual source paths such as `.epub::chapter.xhtml` for traceability.
- Path and stream dispatch, including non-seekable stream support.
- A schema-v5 rich result with chapter pages, HTML blocks, tables, links, forms, bounded manifest assets, metadata, and structured parser diagnostics.
- Chapter-relative, query-only, fragment-only, encoded, root-relative, external, and HTML-base URL projection through the shared EPUB reference contract.
- Structured chapter Markdown for native and ARIA headings, quotes, portable decimal ordered-list markers with preserved starts and value resets, code-language hints, tables, and local EPUB/DPUB-ARIA footnotes.
- Rich `quote`, `code`, and `footnote` blocks, accessible link/image names, exact ordered-list markers, and propagated `officeimo.html.*` capabilities.
- Recoverable diagnostics for unsafe or non-conforming chapter references.

## Boundaries

- Reader adapter configuration belongs here.
- EPUB parsing belongs in `OfficeIMO.Epub`.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.
- Local footnotes become typed Markdown footnotes. Cross-document note targets remain resolved EPUB links.
- Fixed-layout publications are identified and diagnosed; this reader extracts their content and resources but does not emulate a reading-system viewport.
- Unsupported encryption is reported as a security diagnostic and is not decrypted.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Reader.Core`, `OfficeIMO.Reader.Html`, and `OfficeIMO.Epub`; EPUB parsing stays in the native package.
- The EPUB semantics and resource projection layer adds no new package or process dependency.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
