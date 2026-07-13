# OfficeIMO.Epub - EPUB extraction primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Epub)](https://www.nuget.org/packages/OfficeIMO.Epub)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Epub?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Epub)

`OfficeIMO.Epub` provides reusable EPUB extraction primitives for modular OfficeIMO ingestion pipelines.

## Install

```powershell
dotnet add package OfficeIMO.Epub
```

## Quick start

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("book.epub", new EpubReadOptions {
    PreferSpineOrder = true,
    IncludeRawHtml = false,
    MaxChapters = 100
});

Console.WriteLine(book.Title);

foreach (EpubChapter chapter in book.Chapters) {
    Console.WriteLine($"{chapter.Order}. {chapter.Title ?? chapter.Path}");
    Console.WriteLine(chapter.Text);
}

foreach (string warning in book.Warnings) {
    Console.WriteLine(warning);
}
```

### Inspect bounded manifest resources

```csharp
EpubDocument book = EpubDocument.Load("book.epub", new EpubReadOptions {
    IncludeResourceData = true,
    MaxResources = 500,
    MaxResourceBytes = 4L * 1024L * 1024L,
    MaxTotalResourceBytes = 32L * 1024L * 1024L
});

foreach (EpubResource resource in book.Resources) {
    Console.WriteLine($"{resource.Path} ({resource.MediaType}, {resource.LengthBytes} bytes)");
}
```

Manifest metadata is returned even when payload loading is disabled. Payload inclusion is opt-in and bounded per resource, in total, and by resource count; skipped payloads produce warnings.

## What it does

- Opens EPUB files as ZIP containers.
- Parses `META-INF/container.xml` and OPF package metadata.
- Follows OPF manifest and spine ordering.
- Reads nav/NCX labels for chapter titles when available.
- Extracts chapter text from XHTML/XML ASTs.
- Returns deterministic OPF manifest resources with optional bounded payloads.
- Emits extraction warnings for malformed or unreadable content.

## Examples

### Read metadata and spine-ordered chapters

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("handbook.epub", new EpubReadOptions {
    PreferSpineOrder = true,
    IncludeNonLinearSpineItems = false,
    MaxChapters = 50
});

Console.WriteLine(book.Title);
Console.WriteLine(book.Creator);
Console.WriteLine(book.Language);

foreach (var chapter in book.Chapters) {
    Console.WriteLine($"{chapter.Order}. {chapter.Title ?? chapter.Path}");
}
```

### Keep raw chapter HTML when building a converter

```csharp
using OfficeIMO.Epub;

var book = EpubDocument.Load("book.epub", new EpubReadOptions {
    IncludeRawHtml = true,
    MaxChapterBytes = 2L * 1024L * 1024L
});

foreach (var chapter in book.Chapters) {
    File.WriteAllText(
        $"chapter-{chapter.Order:000}.txt",
        chapter.Text);

    if (chapter.Html != null) {
        File.WriteAllText($"chapter-{chapter.Order:000}.xhtml", chapter.Html);
    }
}
```

### Read from a stream and report warnings

```csharp
using OfficeIMO.Epub;

await using var stream = File.OpenRead("upload.epub");
EpubDocument book = EpubDocument.Load(stream, new EpubReadOptions {
    FallbackToHtmlScan = true,
    DeterministicOrder = true
});

foreach (string warning in book.Warnings) {
    Console.WriteLine(warning);
}
```

## Boundaries

- This package owns reusable EPUB parsing primitives.
- Reader integration belongs in `OfficeIMO.Reader.Epub`.
- The parser is read-only and does not attempt CSS layout, scripting, DRM, or package mutation.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
