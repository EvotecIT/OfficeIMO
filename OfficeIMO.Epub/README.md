# OfficeIMO.Epub

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Epub)](https://www.nuget.org/packages/OfficeIMO.Epub)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Epub?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Epub)

Read, create and edit EPUB e-books in .NET. Use `EpubDocument` to extract text and metadata, and `EpubPublication` to create or edit EPUB 2 and 3 packages that survive saving intact.

## Install

```sh
dotnet add package OfficeIMO.Epub
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/Epub/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO.Epub;

EpubPublication publication = EpubPublication.Create("Field guide", "en");
publication.Creator = "Operations";
publication.AddStylesheet("style", "EPUB/style.css",
    "body { font-family: sans-serif; } p { line-height: 1.4; }");
publication.AddChapter("opening", "EPUB/opening.xhtml", "Opening",
    "<h1 id='opening'>Opening</h1><p>A chapter with <em>formatted</em> text.</p>",
    new[] { "style" });
publication.SetNavigation(new[] {
    new EpubNavigationEntry("Opening", "EPUB/opening.xhtml#opening")
});
publication.Save("Field-guide.epub").ThrowIfLoss();

EpubDocument book = EpubDocument.Load("Field-guide.epub");
foreach (EpubChapter chapter in book.Chapters) Console.WriteLine($"{chapter.Order}. {chapter.Title}");
```
<!-- /powerforge:code -->

## What it does

- **Read** publications: container and OPF metadata, spine order, EPUB 3 navigation and EPUB 2 NCX, chapter text and resources, with diagnostics for malformed, encrypted or fixed-layout content.
- **Create** publications with chapters, stylesheets, navigation, metadata and accessibility properties.
- **Import an HTML manuscript** into a structured publication.
- **Edit and preserve** an existing package, including footnotes, bibliographies, indexes, glossaries, book matter and print-page navigation.
- **Preflight and validate** before saving, with save limits and a write report.
- **Before sharing**: inspect and remove hidden data and provenance.

The [EPUB guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.epub-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| Page images | [OfficeIMO.Epub.Image](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Epub.Image/README.md) |

## Limits

EPUB packages are read within size and entry limits, and encrypted resources are reported rather than decrypted. The [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md) lists what is supported, preserved or rejected.

## Learn more

- [OfficeIMO.Epub on officeimo.com](https://officeimo.com/products/epub/) and the [EPUB docs](https://officeimo.com/docs/epub/).
- [EPUB guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.epub-guide.md).
- Reading for search and AI: [OfficeIMO.Reader.Epub](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Reader.Epub/README.md). Command line: `officeimo convert book.epub book.pdf` with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` and `OfficeIMO.Html` (which brings AngleSharp) for content inspection. Container, OPF, navigation and chapter parsing are first-party.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 1 | 0 |
| Read | 2 | 0 | 0 | 0 | 0 | 1 |
| Edit | 1 | 0 | 0 | 1 | 0 | 1 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 4 | 0 | 0 | 0 | 0 | 0 |
| Validate | 2 | 1 | 0 | 0 | 0 | 1 |
| Remove | 2 | 0 | 0 | 0 | 0 | 1 |

The complete rows for `OfficeIMO.Epub` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
