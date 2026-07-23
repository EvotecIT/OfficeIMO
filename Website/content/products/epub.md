---
title: "OfficeIMO.Epub"
description: "Extract EPUB metadata, spine-ordered chapters, navigation, and bounded resources for indexing and ingestion."
layout: product
meta.seo_title: "EPUB extraction for .NET applications"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/products/epub/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/products/epub/" />'
product_color: "#9333ea"
product_label: "EPUB extraction engine"
install: "dotnet add package OfficeIMO.Epub"
nuget: "OfficeIMO.Epub"
docs_url: "/docs/epub/"
---

## EPUB ingestion without a browser engine

`OfficeIMO.Epub` reads the EPUB container, OPF package, navigation, spine, chapters, and resource metadata. It returns deterministic, typed results for search, indexing, migration, and content analysis.

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("handbook.epub", new EpubReadOptions {
    PreferSpineOrder = true,
    MaxChapters = 100
});

foreach (EpubChapter chapter in book.Chapters) {
    Console.WriteLine($"{chapter.Order}: {chapter.Title}");
    Console.WriteLine(chapter.Text);
}
```

## Useful for

- Building a searchable catalog from EPUB title, creator, language, navigation, and chapter text.
- Reading resource metadata without loading every payload into memory.
- Resolving chapter-relative references without network or filesystem access.
- Feeding EPUB into `OfficeIMO.Reader.Epub` for normalized chunks and rich extraction results.
- Rendering EPUB content through focused HTML/Drawing adapters when a visual preview is needed.

The parser is read-only. It does not bypass DRM, execute scripts, perform browser layout, or mutate the package. See the [EPUB extraction guide](/docs/epub/) for bounded resource and Reader examples.
