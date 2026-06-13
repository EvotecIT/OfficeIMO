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

EpubDocument book = EpubReader.Read("book.epub", new EpubReadOptions {
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

## What it does

- Opens EPUB files as ZIP containers.
- Parses `META-INF/container.xml` and OPF package metadata.
- Follows OPF manifest and spine ordering.
- Reads nav/NCX labels for chapter titles when available.
- Extracts chapter text from XHTML/XML ASTs.
- Emits extraction warnings for malformed or unreadable content.

## Boundaries

- This package owns reusable EPUB parsing primitives.
- Reader integration belongs in `OfficeIMO.Reader.Epub`.
- It is still conservative while fuller OPF, spine, and navigation semantics evolve.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
