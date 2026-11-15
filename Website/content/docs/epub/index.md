---
title: Read EPUB Files and Extract Chapters
description: Read EPUB e-books in .NET, get metadata, the table of contents and spine-ordered chapter text, list resources with size limits, and chunk books for search.
order: 41
meta.seo_title: "Read EPUB text and metadata in .NET"
search_terms:
  - epub reader
  - read epub
  - epub to text
  - epub chapters
  - epub metadata
packages:
  - OfficeIMO.Epub
  - OfficeIMO.Reader.Epub
cmdlets:
  - Open-OfficeEpub
  - Get-OfficeEpubChapter
  - Read-OfficeDocument
commands:
  - officeimo reader read
limits: Reading never runs scripts, fetches external resources or removes DRM; encrypted chapters are reported, not decrypted.
verified: 2026-11-14
---

`OfficeIMO.Epub` has two entry points. `EpubDocument` is the read model for extraction: metadata, navigation, chapter text in reading order, and resource listings. `EpubPublication` keeps the whole package so it can be edited and saved (see [Create and edit EPUB files](/docs/epub/create/)). Both open the ZIP container with limits on entry counts, sizes and text volume, and report malformed content as diagnostics instead of failing silently.

```sh
dotnet add package OfficeIMO.Epub
```

## Read metadata and chapter text

Chapters follow the spine order. Titles come from the EPUB 3 navigation document or the EPUB 2 NCX.

### C#

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("handbook.epub", new EpubLoadOptions {
    PreferSpineOrder = true,
    IncludeNonLinearSpineItems = false,
    MaxChapters = 100
});

Console.WriteLine($"{book.Title} by {book.Creator} ({book.Language}, EPUB {book.PackageVersion})");

foreach (EpubChapter chapter in book.Chapters) {
    Console.WriteLine($"{chapter.Order}. {chapter.Title ?? chapter.Path}");
    Console.WriteLine(chapter.Text);
}
```

### PowerShell

```powershell
$book = Open-OfficeEpub -Path .\handbook.epub
$book | Get-OfficeEpubChapter | Select-Object Order, Title, Text
```

### CLI

```sh
officeimo reader read handbook.epub --format markdown --output handbook.md
```

## Walk the table of contents

### C#

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("handbook.epub");

void Print(IReadOnlyList<EpubNavigationItem> items, int depth) {
    foreach (EpubNavigationItem item in items) {
        Console.WriteLine($"{new string(' ', depth * 2)}{item.Label} -> {item.Target}");
        Print(item.Children, depth + 1);
    }
}

Print(book.TableOfContents, 0);
Console.WriteLine($"{book.PageList.Count} print-page markers, {book.Landmarks.Count} landmarks");
```

## List resources without unbounded reads

Resource metadata is always available. Payload bytes are loaded only when you ask for them, and then within per-resource, total and count limits.

### C#

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("book.epub", new EpubLoadOptions {
    IncludeResourceData = true,
    MaxResources = 500,
    MaxResourceBytes = 4L * 1024 * 1024,
    MaxTotalResourceBytes = 32L * 1024 * 1024
});

foreach (EpubResource resource in book.Resources) {
    Console.WriteLine($"{resource.Id}: {resource.Path} ({resource.MediaType}, {resource.LengthBytes} bytes)");
}
```

## Check warnings on uploaded books

### C#

```csharp
using OfficeIMO.Epub;

using FileStream stream = File.OpenRead("upload.epub");
EpubDocument book = EpubDocument.Load(stream, new EpubLoadOptions {
    FallbackToHtmlScan = true
});

foreach (EpubDiagnostic diagnostic in book.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity} {diagnostic.Code} {diagnostic.Path}: {diagnostic.Message}");
}
```

## Split a book into chunks for search

`OfficeIMO.Reader.Epub` produces heading-aware chunks with token budgets for search and retrieval pipelines.

### C#

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

Console.WriteLine($"{result.Chunks.Count} chunks, {result.OutputTokenCount} tokens");
foreach (ReaderChunk chunk in result.Chunks.Take(3)) {
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\handbook.epub
```

### CLI

```sh
officeimo reader read handbook.epub --format json --output handbook.json
```
