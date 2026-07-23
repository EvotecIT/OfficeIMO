---
title: "EPUB Extraction"
description: "Extract EPUB metadata, spine-ordered chapters, navigation, and bounded resources for search and ingestion."
order: 41
meta.seo_title: "Extract EPUB text and metadata in .NET"
---

`OfficeIMO.Epub` is a read-only EPUB engine for ingestion pipelines. It opens the ZIP container, follows OPF manifest and spine order, reads EPUB 3 navigation or EPUB 2 NCX labels, extracts chapter text, and reports malformed or unsupported content.

## Read a book

```shell
dotnet add package OfficeIMO.Epub
```

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("handbook.epub", new EpubReadOptions {
    PreferSpineOrder = true,
    IncludeRawHtml = false,
    MaxChapters = 100
});

Console.WriteLine(book.Title);

foreach (EpubChapter chapter in book.Chapters) {
    Console.WriteLine($"{chapter.Order}. {chapter.Title ?? chapter.Path}");
    Console.WriteLine(chapter.Text);
}
```

## Inspect resources without unbounded expansion

```csharp
EpubDocument book = EpubDocument.Load("book.epub", new EpubReadOptions {
    IncludeResourceData = true,
    MaxResources = 500,
    MaxResourceBytes = 4L * 1024L * 1024L,
    MaxTotalResourceBytes = 32L * 1024L * 1024L
});
```

Resource metadata is available even when payload loading is disabled. Loading resource bytes is opt-in and bounded per resource, in total, and by count.

## Index EPUB with Reader

Install `OfficeIMO.Reader.Epub` when the destination is normalized chunks, rich JSON, or a search/RAG pipeline:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEpubHandler()
    .Build();

ReaderChunkHierarchyResult result = reader.ReadHierarchical(
    "handbook.epub",
    chunkingOptions: new ReaderHierarchicalChunkingOptions {
        MaxTokens = 800,
        OverlapTokens = 80
    });
```

The parser does not execute scripts, fetch external resources, bypass DRM, or mutate the EPUB package. Reference resolution distinguishes container, external, embedded-data, and invalid targets without network or filesystem access.
