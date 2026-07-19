# OfficeIMO.Reader.Core

`OfficeIMO.Reader.Core` is the dependency-light contract and orchestration package for OfficeIMO document ingestion.
It contains normalized result models, limits, deterministic routing, handler registration, processing pipelines, and
capability manifests. It does not reference Word, Excel, PowerPoint, PDF, Email, image, or other format engines.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Core
```

Add only the format packages an application needs:

```powershell
dotnet add package OfficeIMO.Reader.Word
dotnet add package OfficeIMO.Reader.Email
```

Use `OfficeIMO.Reader.All` only when the complete local managed format graph is intentional.

## Build a reader

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .WithMaxConcurrentReads(4)
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("Policy.docx");
```

For dependency-free plain text and an explicit unknown-payload fallback:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPlainTextHandlers()
    .Build();
```

`OfficeDocumentReader.Default` intentionally has no format handlers. This keeps Core honest: adding a format is an
explicit package and builder decision, while every built reader remains immutable and instance-scoped.

## Package selection

| Need | Package |
| --- | --- |
| Contracts, routing, processors, schemas | `OfficeIMO.Reader.Core` |
| Word only | `OfficeIMO.Reader.Word` |
| Excel only | `OfficeIMO.Reader.Excel` |
| PowerPoint only | `OfficeIMO.Reader.PowerPoint` |
| Markdown only | `OfficeIMO.Reader.Markdown` |
| Email artifacts, stores, and OAB | `OfficeIMO.Reader.Email` |
| PDF only | `OfficeIMO.Reader.Pdf` |
| Every local managed handler | `OfficeIMO.Reader.All` |

Other `OfficeIMO.Reader.*` packages follow the same rule: Core plus the format's owning engine. OCR processes,
network clients, hosted providers, and native tools remain explicit host choices and are not composed by All.

## Stable contracts

- `ReaderOptions` and format-neutral input/processing limits
- `ReaderChunk` and the schema-versioned `OfficeDocumentReadResult`
- tables, pages, visuals, assets, links, forms, metadata, and diagnostics
- sync/async path, stream, byte-array, folder, and batch ingestion
- deterministic capability manifests with `OfficeIMO` versus `Custom` handler origins
- bounded nested-content delegation between configured handlers

Public namespaces remain `OfficeIMO.Reader`; the `.Core` name describes package and assembly ownership.
