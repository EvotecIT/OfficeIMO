# OfficeIMO.Reader.Core

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Core)](https://www.nuget.org/packages/OfficeIMO.Reader.Core)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Core?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Core)

The contracts and orchestration behind OfficeIMO's document reader: normalized results, limits, routing, handler registration, processing pipelines and capability manifests. It contains no format engine; you add a reader package for each format you need.

## Install

```sh
dotnet add package OfficeIMO.Reader.Core
dotnet add package OfficeIMO.Reader.Word      # one package per format you read
```

Use [OfficeIMO.Reader.All](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Reader.All/README.md) when you want every local format at once.

## Quick start

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .WithMaxConcurrentReads(4)
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("Policy.docx");
foreach (var item in document.EnumerateContent()) Console.WriteLine(item.Location);
```

`EnumerateContent()` walks paragraphs and tables together in source order, each with its page, slide or sheet location.

## What it does

- **One result shape** for every format: `OfficeDocumentReadResult` with blocks, tables, pages, visuals, assets, links, forms, metadata, diagnostics and chunks with source locations.
- **Read from anywhere**: paths, streams, byte arrays, folders and batches, synchronously or asynchronously, with incremental ingestion and batches bounded by source bytes.
- **Limits and budgets** for input size, nesting and decoding, and immutable read policies you can reuse.
- **Find content by page**, inspect review content such as comments and tracked changes, and detect encryption without opening a file.
- **Export tables** as CSV, Markdown or JSON, and project text and metadata for search indexes.
- **Custom handlers** for your own formats, with capability manifests that tell OfficeIMO handlers from custom ones.

The [Reader guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.reader.md) shows each of these with examples.

## Reader packages

| Need | Package |
| --- | --- |
| Word, Excel, PowerPoint | `OfficeIMO.Reader.Word`, `OfficeIMO.Reader.Excel`, `OfficeIMO.Reader.PowerPoint` |
| PDF | `OfficeIMO.Reader.Pdf` |
| Markdown, HTML | `OfficeIMO.Reader.Markdown`, `OfficeIMO.Reader.Html` |
| Email artifacts, stores and address books | `OfficeIMO.Reader.Email` |
| Scanned pages and images | `OfficeIMO.Reader.Ocr` with an OCR provider |
| Every local managed handler | `OfficeIMO.Reader.All` |
| `Microsoft.Extensions.DocumentExtraction` provider (preview) | `OfficeIMO.Reader.DocumentExtraction` |

Every other `OfficeIMO.Reader.*` package follows the same rule: Core plus that format's engine. OCR processes, network clients and hosted providers stay explicit choices.

## Limits

Extracted text follows each format's logical order, which isn't always the original reading order, and scanned pages need an OCR provider. The [content safety matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.content-safety-support-matrix.md) covers how hidden and prompt-like content is reported.

## Learn more

- [Reader guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.reader.md) and the [Reader docs](https://officeimo.com/docs/reader/) on officeimo.com.
- Command line and agents: `officeimo` extraction and search with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `System.Text.Json` only; no format engine. Public types are in the `OfficeIMO.Reader` namespace.
- License: MIT.

