---
title: "Publisher Publications"
description: "Read Publisher .pub files into pages, stories, tables and images, apply supported edits, and export SVG, PDF and Reader output in .NET."
order: 56
meta.seo_title: "Read and convert Microsoft Publisher files in .NET"
---

`OfficeIMO.Publisher` reads Microsoft Publisher publications into positioned pages, styled text stories, native table grids and embedded image assets. It applies supported edits to loaded publications, authors new publications in memory, and renders the current model to SVG. Publisher does not need to be installed.

Two companion packages build on the same model:

| Package | Purpose |
|---|---|
| `OfficeIMO.Publisher` | Native reading, supported source edits, in-memory authoring, SVG rendering |
| `OfficeIMO.Publisher.Pdf` | Multi-page PDF output through the shared drawing and PDF engines |
| `OfficeIMO.Reader.Publisher` | Stories, tables, page inventory and image descriptors in the shared Reader result |

## Install

```shell
dotnet add package OfficeIMO.Publisher
dotnet add package OfficeIMO.Publisher.Pdf
```

## Read a publication

```csharp
using OfficeIMO.Publisher;

using var publication = PublisherDocument.Load("newsletter.pub");

foreach (var page in publication.Pages)
    Console.WriteLine($"Page {page.Id}: {page.Width} x {page.Height} points");

foreach (var story in publication.TextStories)
    Console.WriteLine(story.Text);

foreach (var diagnostic in publication.ReadReport.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

`Pages` follows publication order. `MasterPages` holds the recovered master definitions, whose content is also applied to the pages that use them. Coordinates are in points with a top-left origin. `TextStories` keeps complete story text, including content that overflows its last frame. Each page also exposes `TextFrames` and `Tables` with native identifiers, story ranges, column settings, cell spans and cell text.

The read report separates approximations, omissions and unassessed content. A successful load is not a claim of complete Publisher fidelity; call `ReadReport.ThrowIfLoss()` when any loss must stop processing.

## Convert to SVG and PDF

```csharp
using OfficeIMO;
using OfficeIMO.Publisher;
using OfficeIMO.Publisher.Pdf;

using var publication = PublisherDocument.Load("brochure.pub");

string svg = publication.ToSvg(0, null, out OfficeConversionReport svgReport);
File.WriteAllText("page-1.svg", svg);

OfficeConversionReport pdfReport = publication.SaveAsPdf("brochure.pdf");
foreach (var diagnostic in pdfReport.Diagnostics)
    Console.WriteLine($"{diagnostic.LossKind}: {diagnostic.Message}");
```

PDF pages keep publication order and physical dimensions; masters are applied rather than emitted as extra pages. Supply fonts through `PublisherLoadOptions.Fonts` so story flow, measurement and embedded glyphs use the same faces. Set `FailOnLoss` on `PublisherPdfExportOptions` to reject reported loss before anything is written.

## Edit and save a loaded publication

```csharp
using OfficeIMO.Publisher;

using var publication = PublisherDocument.Load("flyer.pub");
var page = publication.Pages[0];
publication.SetPageObjectOrder(page.Id, page.ObjectIds.Reverse().ToArray());

PublisherSaveReport report = publication.Save("flyer-reordered.pub",
    new PublisherSaveOptions { AllowLoss = true });
foreach (var diagnostic in report.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

Saving an unchanged publication produces a byte-exact copy of the source. Supported edits include page object order, frame, shape and picture bounds, existing solid paint, frame columns and insets, story text and paragraph formatting, picture replacement with PNG or JPEG, existing table cell text, and assigning existing masters. Edits outside the documented profiles refuse native output instead of writing a damaged file.

The default loss policy blocks edited output; set `AllowLoss = true` to accept the reported fidelity. File saves require a `.pub` extension and refuse to overwrite unless `FileConflictPolicy` is set to `Replace`.

## Author a publication in memory

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher;

using var newsletter = PublisherDocument.Create(OfficeLength.Points(792), OfficeLength.Points(612));
var firstPage = newsletter.Pages[0];
var header = newsletter.AddMasterPage("Newsletter header");
newsletter.AddTextFrame(header.Id, OfficeLength.Points(36), OfficeLength.Points(18),
    OfficeLength.Points(600), OfficeLength.Points(30), new[] {
    new OfficeRichTextParagraph(new[] {
        new OfficeRichTextRun("Weekly publication", 16, OfficeColor.Black, bold: true, fontFamily: "Arial")
    })
});
newsletter.SetPageMaster(firstPage.Id, header.Id);
newsletter.AddTextFrame(firstPage.Id, OfficeLength.Points(36), OfficeLength.Points(84),
    OfficeLength.Points(300), OfficeLength.Points(160), new[] {
    new OfficeRichTextParagraph(new[] {
        new OfficeRichTextRun("Stories, masters and pictures render to SVG and PDF.", 12, OfficeColor.Black, fontFamily: "Arial")
    })
});
File.WriteAllText("newsletter-page.svg", newsletter.ToSvg());
```

In-memory publications support pages, single and left/right master pages, linked text frames, styled text, embedded PNG or JPEG pictures, tables and primitive shapes. They render to SVG, PDF and Reader output. Saving a new publication as a native `.pub` file is not supported; native saving requires a loaded source.

## Read through OfficeIMO.Reader

```csharp
using OfficeIMO.Publisher;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Publisher;

using var publication = PublisherDocument.Load("newsletter.pub");
OfficeDocumentReadResult result = publication.ToOfficeDocumentReadResult(
    sourceName: "newsletter.pub",
    readerOptions: new ReaderOptions { MaxChars = 4_000 });

foreach (OfficeDocumentBlock block in result.Blocks)
    Console.WriteLine(block.Text);
```

The Reader projection keeps every story once in native order, native table rows, bullet markers and image descriptors. Page artwork, frame flow and typography are omitted from this semantic view and reported with `PUB_READER_LAYOUT_OMITTED`. Register the handler with `AddPublisherHandler` on `OfficeDocumentReaderBuilder` to read `.pub` files alongside other formats.

## Limits

- The Publisher 2002-and-later Contents/Quill generation is qualified. Publisher 97/98 and 2000 files are recognized and rejected; encrypted or damaged files are not decrypted or salvaged.
- Native WMF and EMF pictures stay available in `Images`; page rendering shows a placeholder unless an image codec is supplied.
- Rendering uses OfficeIMO's managed text layout. Line breaks, font substitution and some native effects can differ from Publisher and are reported as approximations.
- Dynamic fields keep their cached text and are not evaluated. Macros, external resource refresh and embedded-object activation are outside the contract.
- Page resizing, deletion and reordering are not provided.

## Related

- [Publisher support contract](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Publisher/SUPPORT.md)
- [Publisher API reference](/api/publisher/) and [Publisher PDF API reference](/api/publisher-pdf/)
- [Reader guide](/docs/reader/)
- [OfficeIMO.Publisher product page](/products/publisher/)
