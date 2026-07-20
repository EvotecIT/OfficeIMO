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

## Find content by page

Page-aware reading stays on `OfficeDocumentReadResult`; it is not a separate conversion path. A format adapter
populates `Pages`, and Reader Core provides shared location, search, and page-scoped Markdown helpers:

```csharp
OfficeDocumentSearchResult matches = document.Search("retention period");

foreach (OfficeDocumentSearchHit hit in matches.Hits) {
    Console.WriteLine(hit.Block.Text);
    foreach (OfficeDocumentPageLocation location in hit.Pages) {
        Console.WriteLine(location.Display); // for example: Page 5 of 20
    }
}

string pageMarkedMarkdown = document.ToPageMarkedMarkdown();
```

Page boundaries are not equally authoritative in every source format:

| Format | Page provenance | Reader behavior |
| --- | --- | --- |
| PDF | `Native` | Uses fixed pages and source geometry from the PDF logical model. |
| Word | `Computed` | Opt-in best-effort pagination through the OfficeIMO.Word layout engine. |
| RTF | `ExplicitBreak` | Opt-in reconstruction from explicit/saved page and section-break hints; automatic overflow is not calculated. |

Use `document.GetPageProvenance()` when page accuracy affects citations. `GetPageMarkdown()` returns separate
page values, while `ToPageMarkedMarkdown()` produces one portable Markdown string with HTML page markers.
The original document-wide `Markdown`, `Blocks`, and `Chunks` remain available on the same result.

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
