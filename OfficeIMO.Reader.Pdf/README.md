# OfficeIMO.Reader.Pdf - PDF reader and compatibility projection bridge

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Pdf)](https://www.nuget.org/packages/OfficeIMO.Reader.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Pdf)

`OfficeIMO.Reader.Pdf` reads PDF artifacts into the normalized Reader model. For existing Reader workflows it also contains a thin compatibility bridge that maps an `OfficeDocumentReadResult` into the dependency-free `OfficeDocumentModel`; the actual output policy and PDF composition are owned by `OfficeIMO.Pdf`. It does not pull Word, Excel, PowerPoint, Email, or the all-adapters composition package.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Pdf
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler()
    .Build();

IReadOnlyList<ReaderChunk> chunks = reader
    .Read("invoice.pdf")
    .ToList();
```

## Examples

### Read page-aware Markdown chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler()
    .Build();

foreach (var chunk in reader.Read("manual.pdf")) {
    Console.WriteLine($"Page {chunk.Location.Page}: {chunk.Id}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}

OfficeDocumentReadResult document = reader.ReadDocument("manual.pdf");
OfficeDocumentSearchResult matches = document.Search("installation");
Console.WriteLine(string.Join(", ", matches.Hits
    .SelectMany(hit => hit.Pages)
    .Select(page => page.Display)));
```

Cross-page paragraph continuation evidence is projected into document metadata
without rewriting page-local text or chunk boundaries. The default confidence
policy is conservative and can be tightened at registration time:

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler(new ReaderPdfOptions {
        ParagraphContinuationOptions = new PdfLogicalParagraphContinuationOptions {
            MinimumConfidence = 0.85,
            RejoinLineEndingHyphens = true
        }
    })
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("manual.pdf");
IReadOnlyList<OfficeDocumentMetadataEntry> continuations = document.Metadata
    .Where(entry => entry.Category == "pdf.paragraph.continuation")
    .ToArray();
```

The semantic reader processes at most 1,000 selected pages by default. Set the
semantic pipeline limit explicitly when a larger document is expected:

```csharp
OfficeDocumentReader largePdfReader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler(new ReaderPdfOptions {
        ReadOptions = new PdfReadOptions {
            Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 2_500 }
        }
    })
    .Build();
```

`ReaderPdfOptions.ReadOptions` is the canonical PDF semantic configuration. Use
it for `Profile`, `PageSelection`, `LayoutOptions`, custom stages, and semantic
work limits; the Reader adapter does not maintain a second set of PDF parsing
options.

### Read a stream with input limits

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler()
    .Build();

using var stream = File.OpenRead("large-report.pdf");
var chunks = reader.Read(stream, "large-report.pdf", new ReaderOptions {
    MaxChars = 12_000,
    MaxInputBytes = 100L * 1024L * 1024L
}).ToList();

foreach (var chunk in chunks.Where(chunk => chunk.Diagnostics != null)) {
    Console.WriteLine($"{chunk.Id}: {chunk.Diagnostics!.TableCount} table(s)");
}
```

### Project a normalized document into PDF

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;

OfficeDocumentReadResult normalized = reader.ReadDocument("manual.pdf");
PdfDocumentConversionResult conversion = normalized.ToPdfDocumentResult(
    new PdfProjectionOptions {
        PagePolicy = PdfProjectionPagePolicy.PreserveSourcePages,
        AssetPolicy = PdfProjectionAssetPolicy.EmbedSupportedImages,
        RasterDecodeOptions = new OfficeRasterDecodeOptions {
            AnimationPolicy = OfficeRasterAnimationPolicy.RejectAnimated
        },
        LinkPolicy = PdfProjectionLinkPolicy.PreserveUriLinks,
        FormPolicy = PdfProjectionFormPolicy.RenderCurrentValues
    });

PdfSaveResult save = conversion.Save("normalized-manual.pdf");
foreach (PdfConversionWarning warning in save.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The same projection accepts normalized results produced by Email, EPUB, Visio,
Office, and other Reader adapters without taking dependencies on those packages.
Source diagnostics are merged with PDF generation diagnostics. Email attachment,
EPUB resource/pagination, and Visio preview/semantic fallback decisions are
recorded explicitly. Page-scoped and document-scoped normalized collections are
reconciled by the neutral model identities, so aggregate resources are retained
without repeating page content. The Drawing raster policy is used for selected
GIF frames and animation rejection; selecting a frame emits loss evidence rather
than silently flattening animation. This does not advertise a direct format
converter until that route has its own artifact and degradation gates.

### Register alongside other ingestion adapters

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Zip;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddHtmlHandler()
    .AddPdfHandler()
    .AddZipHandler()
    .Build();

var chunks = reader.ReadFolder("KnowledgeBase", new ReaderFolderOptions {
    Recurse = true,
    DeterministicOrder = true,
    MaxFiles = 500
}).ToList();
```

## What it emits

- Page-aware chunks with `ReaderLocation.Page`.
- Native fixed-page membership and geometry compatible with Reader Core location, search, and page-Markdown helpers.
- Markdown text, logical tables, column profiles, table diagnostics, and confidence signals.
- Conservative paragraph continuation metadata with page range, confidence, evidence, segment count, and rejoined-hyphen count.
- Source/security/form/catalog-metadata/open-action/active-content counters in `ReaderChunk.Diagnostics`.
- Document metadata for XMP, output intents, tagged structure, optional content/layers, attachments, security/signatures, navigation, links, forms, annotations, and passive actions.
- Passive action summaries without executable payloads.
- Image placeholders and visual geometry when available.
- Link annotations and typed form fields when available.
- A source-neutral normalized-document-to-PDF projection with page, asset, link,
  and form policies plus merged conversion evidence.

## Boundaries

- Reader adapter configuration belongs here.
- PDF parsing, logical readback, generation, and safety behavior belongs in `OfficeIMO.Pdf`.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.
- This package maps between those two owners; source-format packages continue to
  own their normalization rules and do not gain a second PDF engine here.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Reader.Core`, `OfficeIMO.Core`, and the first-party `OfficeIMO.Pdf` engine own normalized content, raster compatibility, PDF parsing/generation, assets, and diagnostics.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
