# OfficeIMO.Reader.Html - HTML reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Html)](https://www.nuget.org/packages/OfficeIMO.Reader.Html)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Html?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Html)

`OfficeIMO.Reader.Html` registers a modular HTML ingestion adapter for `OfficeIMO.Reader`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Html
```

## Register

```csharp
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();
```

For untrusted or size-sensitive HTML:

```csharp
DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(
    htmlOptions: ReaderHtmlOptions.CreateUntrustedHtmlProfile(maxInputCharacters: 100_000),
    replaceExisting: true);
```

## What it emits

- HTML converted to Markdown through `OfficeIMO.Markdown.Html`.
- Markdown-shaped `ReaderChunk` output.
- Table extraction with `ReaderTable.ColumnProfiles`.
- Heading-aware chunk metadata when `ReaderOptions.MarkdownChunkByHeadings` is enabled.
- HTML-to-Markdown profile, transform, converter, and visual round-trip option pass-through.

## Boundaries

- Reader adapter registration belongs here.
- HTML to Markdown conversion belongs in `OfficeIMO.Markdown.Html`.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
