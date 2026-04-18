---
title: "OfficeIMO.Reader"
description: "Unified document extraction from Word, Excel, PowerPoint, Markdown, PDF, and text-like inputs. Deterministic chunks for indexing and AI workflows."
layout: product
product_color: "#4f46e5"
install: "dotnet add package OfficeIMO.Reader"
nuget: "OfficeIMO.Reader"
docs_url: "/docs/reader/"
api_url: "/api/reader/"
preview_id: "reader"
---

## Why OfficeIMO.Reader?

OfficeIMO.Reader provides a single API to extract structured content from common document formats. Feed it Word, Excel, PowerPoint, Markdown, PDF, or text-like inputs and get back normalized chunks with location data, source hashes, and token estimates. Optional adapters extend the same reader model to CSV/TSV, JSON, XML, HTML, EPUB, and ZIP-oriented ingestion. It is purpose-built for RAG pipelines, search indexing, and any workflow where you need reproducible document slices instead of ad-hoc parsers.

## Features

- **Extract from Word, Excel, PowerPoint, Markdown & PDF** -- one API for all major document formats
- **Deterministic extraction chunks** -- emit stable chunk boundaries with configurable character and row limits
- **Heading-aware extraction with citations** -- preserve document structure and source locations on each chunk
- **Token estimates per chunk** -- budget prompts and indexing payloads without an extra preprocessing pass
- **Folder batch processing** -- process entire directories with progress callbacks, skip reporting, and cancellation support
- **Pluggable handler registration** -- register custom extractors for proprietary or domain-specific formats
- **Adapter packages** -- add CSV, JSON, XML, HTML, EPUB, ZIP, or structured text support without bloating every reader deployment

## What teams build with Reader

| Workflow | Output | Why Reader fits |
|----------|--------|-----------------|
| Knowledge ingestion services | Chunked text plus source IDs and token estimates for vector stores and semantic search | One extractor handles mixed Office and adjacent formats with the same result model |
| Compliance and review pipelines | Searchable evidence bundles with headings and citations | Stable chunk boundaries make reviews and re-runs easier to compare |
| File-share indexing jobs | Normalized documents ready for Lucene, Elasticsearch, or Azure AI Search | Batch extraction works well in workers, scheduled jobs, and containers |
| Content migration tools | Markdown, JSON, or sidecar artifacts derived from legacy documents | Structured extraction keeps enough source context to transform before re-emitting |

## Reader package family

| Package | Use it when |
|---------|-------------|
| `OfficeIMO.Reader` | You need the core mixed Office, Markdown, and PDF extraction facade. |
| `OfficeIMO.Reader.Csv` | CSV and TSV files should flow through the same chunking and metadata model. |
| `OfficeIMO.Reader.Json` | JSON payloads need deterministic document slices for indexing or review. |
| `OfficeIMO.Reader.Xml` | XML documents should keep element-aware source context during extraction. |
| `OfficeIMO.Reader.Text` | You want the combined CSV, JSON, and XML adapter path in one package. |
| `OfficeIMO.Reader.Html` | HTML input should be normalized through the Markdown/Reader pipeline. |
| `OfficeIMO.Reader.Epub` | EPUB books or packaged publications belong in the same ingestion workflow. |
| `OfficeIMO.Reader.Zip` | ZIP archives need safe traversal and chunking as part of a reader job. |

## Quick start

```csharp
using OfficeIMO.Reader;

var chunks = DocumentReader.Read("report.docx", new ReaderOptions
{
    MaxChars = 4_000,
    IncludeWordFootnotes = true,
    ComputeHashes = true
}).ToList();

foreach (var chunk in chunks)
{
    Console.WriteLine($"{chunk.Id} :: {chunk.Kind} :: ~{chunk.TokenEstimate ?? 0} tokens");
    Console.WriteLine(chunk.Location.HeadingPath ?? chunk.Location.Path ?? "unknown");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
    Console.WriteLine();
}

var documents = DocumentReader.ReadFolderDocuments(
    folderPath: "./documents",
    folderOptions: new ReaderFolderOptions
    {
        Recurse = true,
        DeterministicOrder = true,
        MaxFiles = 500
    },
    options: new ReaderOptions
    {
        MaxChars = 4_000,
        ComputeHashes = true
    },
    onProgress: progress =>
        Console.WriteLine($"{progress.Kind}: scanned={progress.FilesScanned}, parsed={progress.FilesParsed}, skipped={progress.FilesSkipped}, chunks={progress.ChunksProduced}")
).ToList();

Console.WriteLine($"Processed {documents.Count} files");
Console.WriteLine($"Parsed {documents.Count(d => d.Parsed)} files");
Console.WriteLine($"Returned {documents.Sum(d => d.ChunksProduced)} chunks");
```

## Typical ingestion flow

1. Detect the source format and extract chunks, headings, tables, visuals, and source information with one API call.
2. Normalize the result into a shape your pipeline understands, regardless of whether the input was Word, Excel, PowerPoint, Markdown, PDF, or text.
3. Tune `ReaderOptions` so citations and downstream embeddings stay stable across repeated runs.
4. Store the chunks, source hashes, and source references in your vector store, search index, or audit trail.
5. Reuse the same extractor in local tools, hosted services, or CI jobs without changing the document pipeline.

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.Reader targets the same cross-platform .NET runtimes as the packages it builds on, and the core extraction flow is a good fit for containers, hosted services, and server-side indexing jobs. As with any mixed-format pipeline, validate your exact deployment shape, input set, and runtime targets before treating it as broadly portable across every environment.

## Related guides

| Guide | Description |
|-------|-------------|
| [Reader documentation](/docs/reader/) | Learn the core extraction model, chunking workflow, adapter family, and ingestion patterns. |
| [Reader API reference](/api/reader/) | Browse `DocumentReader`, `ReaderOptions`, chunk models, and handler extension points. |
| [AOT and trimming](/docs/advanced/aot-trimming/) | Review runtime and deployment guidance for lean extraction services. |
| [Reader tutorial](/blog/reading-documents-with-reader/) | Walk through chunk inspection, folder ingestion, and indexing-oriented extraction patterns. |
| [OfficeIMO.Markdown](/products/markdown/) | Pair extraction with markdown rendering and transformation workflows. |
| [Downloads](/downloads/) | Pick the core reader package or one of the specialized adapter packages. |
