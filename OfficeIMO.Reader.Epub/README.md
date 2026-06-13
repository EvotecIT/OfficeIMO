# OfficeIMO.Reader.Epub - EPUB reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Epub)](https://www.nuget.org/packages/OfficeIMO.Reader.Epub)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Epub?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Epub)

`OfficeIMO.Reader.Epub` bridges `OfficeIMO.Epub` output into `OfficeIMO.Reader` chunk contracts.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Epub
```

## Register

```csharp
using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;

DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(new EpubReadOptions {
    PreferSpineOrder = true,
    IncludeRawHtml = false,
    MaxChapters = 100
});
```

## Examples

### Read chapters as Reader chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;

DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler();

foreach (var chunk in DocumentReader.Read("book.epub", new ReaderOptions {
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

DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(new EpubReadOptions {
    FallbackToHtmlScan = true,
    MaxChapterBytes = 2L * 1024L * 1024L
}, replaceExisting: true);

await using var stream = File.OpenRead("upload.epub");
var chunks = DocumentReader.Read(stream, "upload.epub").ToList();

foreach (string warning in chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>())) {
    Console.WriteLine(warning);
}
```

## What it emits

- Chapter-to-chunk projection.
- Max-character chunk splitting.
- Markdown and text chunk payloads.
- Warning chunks propagated from EPUB parser warnings.
- Virtual source paths such as `.epub::chapter.xhtml` for traceability.
- Path and stream dispatch, including non-seekable stream support.

## Boundaries

- Reader adapter registration belongs here.
- EPUB parsing belongs in `OfficeIMO.Epub`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
