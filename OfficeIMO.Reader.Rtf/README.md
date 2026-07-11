# OfficeIMO.Reader.Rtf - RTF reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Rtf)](https://www.nuget.org/packages/OfficeIMO.Reader.Rtf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Rtf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Rtf)

`OfficeIMO.Reader.Rtf` registers a bounded RTF adapter for `OfficeIMO.Reader` using the shared `OfficeIMO.Rtf` semantic model.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Rtf
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;

DocumentReaderRtfRegistrationExtensions.RegisterRtfHandler();

IReadOnlyList<ReaderChunk> chunks = DocumentReader
    .Read("clinical-note.rtf")
    .ToList();
```

## Examples

### Read block-aware chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;

DocumentReaderRtfRegistrationExtensions.RegisterRtfHandler();

foreach (var chunk in DocumentReader.Read("policy.rtf")) {
    Console.WriteLine($"{chunk.Id}: {chunk.Location.SourceBlockKind}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read a stream with input limits

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;

DocumentReaderRtfRegistrationExtensions.RegisterRtfHandler();

using var stream = File.OpenRead("large-note.rtf");
var rtfOptions = new ReaderRtfOptions(); // bounded core profile by default
var chunks = DocumentReaderRtfExtensions.ReadRtf(stream, "large-note.rtf", new ReaderOptions {
    MaxChars = 12_000,
    MaxInputBytes = 100L * 1024L * 1024L
}, rtfOptions).ToList();

rtfOptions.ConversionReport.RequireNoLoss();
```

### Read the semantic RTF envelope

```csharp
OfficeDocumentReadResult document =
    DocumentReaderRtfExtensions.ReadRtfDocumentResult("form.rtf");

foreach (OfficeDocumentFormField field in document.Forms) {
    Console.WriteLine($"{field.Name}: {field.Value}");
}

foreach (OfficeDocumentAsset asset in document.Assets) {
    Console.WriteLine($"{asset.Kind}: {asset.FileName}");
}
```

After registration, `DocumentReader.ReadDocument("form.rtf")` uses the same native rich mapping.

## What it emits

- RTF input kind and deterministic source/chunk metadata.
- Paragraph, list, table, note, header/footer, object, shape, and image placeholder chunks from the semantic RTF model.
- Markdown-friendly table output plus `ReaderTable` payloads for parsed RTF tables.
- Parser and binder diagnostics as reader warnings when requested.
- A shared conversion report for flattened, omitted, and blocked features.
- A schema-v5 rich result containing semantic blocks, tables, hyperlinks, form fields, image visuals and payloads, embedded-object assets, metadata, and structured diagnostics.

## Boundaries

- Reader adapter registration belongs here.
- RTF parsing, syntax preservation, semantic binding, and writing belong in `OfficeIMO.Rtf`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
