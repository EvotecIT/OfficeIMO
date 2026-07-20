# OfficeIMO.Reader.Rtf - RTF reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Rtf)](https://www.nuget.org/packages/OfficeIMO.Reader.Rtf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Rtf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Rtf)

`OfficeIMO.Reader.Rtf` provides a bounded RTF adapter for `OfficeIMO.Reader.Core` using the shared `OfficeIMO.Rtf` semantic model.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Rtf
```

## Configure

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddRtfHandler()
    .Build();

IReadOnlyList<ReaderChunk> chunks = reader
    .Read("clinical-note.rtf")
    .ToList();
```

## Examples

### Read block-aware chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddRtfHandler()
    .Build();

foreach (var chunk in reader.Read("policy.rtf")) {
    Console.WriteLine($"{chunk.Id}: {chunk.Location.SourceBlockKind}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read a stream with input limits

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;

var rtfOptions = new ReaderRtfOptions(); // bounded core profile by default
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddRtfHandler(rtfOptions)
    .Build();

using var stream = File.OpenRead("large-note.rtf");
RtfConversionResult<IReadOnlyList<ReaderChunk>> result =
    RtfReaderAdapter.ReadResult(stream, "large-note.rtf", new ReaderOptions {
    MaxChars = 12_000,
    MaxInputBytes = 100L * 1024L * 1024L
}, rtfOptions);

IReadOnlyList<ReaderChunk> chunks = result.RequireNoLoss();
```

### Read the semantic RTF envelope

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddRtfHandler()
    .Build();
OfficeDocumentReadResult document = reader.ReadDocument("form.rtf");

foreach (OfficeDocumentFormField field in document.Forms) {
    Console.WriteLine($"{field.Name}: {field.Value}");
}

foreach (OfficeDocumentAsset asset in document.Assets) {
    Console.WriteLine($"{asset.Kind}: {asset.FileName}");
}
```

`reader.ReadDocument("form.rtf")` uses the same native rich mapping as the adapter-specific APIs.

### Reconstruct explicit page locations

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddRtfHandler(new ReaderRtfOptions {
        IncludePageLocations = true
    })
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("policy.rtf");
OfficeDocumentSearchResult matches = document.Search("retention period");

foreach (OfficeDocumentSearchHit hit in matches.Hits) {
    Console.WriteLine(string.Join(", ", hit.Pages.Select(page => page.Display)));
}
```

RTF is flow content rather than a fixed-page format. This option reconstructs membership from `\page`,
`\softpage`, `\pagebb`, and page-starting section controls. A saved `\nofpages` value contributes the total
page count when present. Reader emits `OfficeDocumentPageProvenance.ExplicitBreak` and an informational diagnostic;
it does not estimate page changes caused only by automatic text overflow.

## What it emits

- RTF input kind and deterministic source/chunk metadata.
- Paragraph, list, table, note, header/footer, object, shape, and image placeholder chunks from the semantic RTF model.
- Markdown-friendly table output plus `ReaderTable` payloads for parsed RTF tables.
- Parser and binder diagnostics as reader warnings when requested.
- A shared conversion report for flattened, omitted, and blocked features.
- A schema-v5 rich result containing semantic blocks, tables, hyperlinks, form fields, image visuals and payloads, embedded-object assets, metadata, and structured diagnostics.
- Optional explicit-break page membership compatible with Reader Core location, search, and page-Markdown helpers.

## Boundaries

- Reader adapter configuration belongs here.
- RTF parsing, syntax preservation, semantic binding, and writing belong in `OfficeIMO.Rtf`.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and the first-party `OfficeIMO.Rtf` engine own parsing, semantic binding, chunks, assets, and diagnostics.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
