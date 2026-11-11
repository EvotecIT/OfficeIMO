---
title: "OfficeIMO.Publisher"
description: "Read Microsoft Publisher .pub files into pages, stories, tables and images, apply supported edits, and export SVG and PDF without Publisher."
layout: product
meta.seo_title: "Microsoft Publisher reading and conversion for .NET"
product_color: "#077568"
product_label: "Microsoft Publisher engine"
install: "dotnet add package OfficeIMO.Publisher"
nuget: "OfficeIMO.Publisher"
docs_url: "/docs/publisher/"
api_url: "/api/publisher/"
meta.software.name: "OfficeIMO.Publisher"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.nuget.org/packages/OfficeIMO.Publisher"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## Open Publisher files on any platform

`OfficeIMO.Publisher` decodes Publisher 2002-and-later publications into positioned pages, master pages, styled text stories, native tables and embedded images. Every load carries a recovery report that names approximations, omissions and unassessed content.

```csharp
using OfficeIMO;
using OfficeIMO.Publisher;
using OfficeIMO.Publisher.Pdf;

using var publication = PublisherDocument.Load("brochure.pub");
foreach (var story in publication.TextStories)
    Console.WriteLine(story.Text);

OfficeConversionReport report = publication.SaveAsPdf("brochure.pdf");
Console.WriteLine($"Reported loss: {report.HasLoss}");
```

## Useful for

- Archiving legacy newsletters, flyers and brochures as PDF or SVG.
- Extracting complete story text and table rows for search and indexing.
- Updating text, layout, pictures and table cells in existing publications with an explicit fidelity report.
- Composing publications in memory with masters, linked frames, pictures and tables, then rendering them.

## Packages

- `OfficeIMO.Publisher`: native reading, supported source edits, in-memory authoring and SVG.
- `OfficeIMO.Publisher.Pdf`: multi-page PDF through the shared OfficeIMO PDF engine.
- `OfficeIMO.Reader.Publisher`: stories, tables, page inventory and image descriptors for OfficeIMO.Reader.

## Current limits

Publisher 97/98 and 2000 files are rejected, and encrypted or damaged files are not salvaged. Unchanged publications save as exact copies; edits are saved only within documented profiles. New publications render to SVG and PDF but cannot be saved as native `.pub` files. Rendering uses managed layout, so line breaks can differ from Publisher.

Start with the [Publisher guide](/docs/publisher/) or review the [support contract](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Publisher/SUPPORT.md).
