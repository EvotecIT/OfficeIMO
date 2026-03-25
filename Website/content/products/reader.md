---
title: "OfficeIMO.Reader"
description: "Unified document extraction from Word, Excel, PowerPoint, Markdown, PDF, and more. Deterministic chunking for AI workflows."
layout: product
product_color: "#4f46e5"
install: "dotnet add package OfficeIMO.Reader"
nuget: "OfficeIMO.Reader"
docs_url: "/docs/reader/"
api_url: ""
---

## Why OfficeIMO.Reader?

OfficeIMO.Reader provides a single API to extract structured content from any common document format. Feed it Word, Excel, PowerPoint, Markdown, or PDF files and get back clean, chunked text with metadata. It is purpose-built for RAG pipelines, search indexing, and any workflow where you need to turn documents into structured data.

## Features

- **Extract from Word, Excel, PowerPoint, Markdown & PDF** -- one API for all major document formats
- **Deterministic chunking for AI** -- split documents into reproducible, overlapping chunks with configurable size and stride
- **Heading-aware extraction with citations** -- preserve document structure and emit source citations for each chunk
- **Token estimation** -- estimate token counts for OpenAI, Anthropic, and other LLM tokenizers without external calls
- **Folder batch processing** -- process entire directories with progress tracking and cancellation support
- **Pluggable handler registration** -- register custom extractors for proprietary or domain-specific formats

## Quick start

```csharp
using OfficeIMO.Reader;

// Extract from a single file
var result = DocumentReader.Extract("report.docx");
Console.WriteLine($"Title: {result.Title}");
Console.WriteLine($"Pages: {result.PageCount}");
Console.WriteLine($"Text length: {result.Text.Length}");

// Chunk for AI ingestion
var chunks = DocumentReader.Chunk("report.docx", new ChunkOptions
{
    MaxTokens = 512,
    Overlap = 64,
    PreserveHeadings = true
});

foreach (var chunk in chunks)
{
    Console.WriteLine($"[{chunk.Index}] {chunk.Heading} -- {chunk.EstimatedTokens} tokens");
    Console.WriteLine(chunk.Text);
    Console.WriteLine($"  Source: {chunk.Citation}");
    Console.WriteLine();
}

// Batch process a folder
var results = await DocumentReader.ExtractFolderAsync("./documents/",
    pattern: "*.docx;*.xlsx;*.pptx;*.pdf",
    progress: new Progress<ExtractionProgress>(p =>
        Console.WriteLine($"{p.FileName} -- {p.Status}")
    ));

Console.WriteLine($"Processed {results.Count} documents");
Console.WriteLine($"Total chunks: {results.Sum(r => r.Chunks.Count)}");
```

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.Reader runs on Windows, Linux, and macOS. It has no native dependencies and works in containers, Azure Functions, AWS Lambda, and any .NET hosting environment.

## Related guides

| Guide | Description |
|-------|-------------|
| [Reader documentation](/docs/reader/) | Learn the core extraction model, chunking workflow, and ingestion patterns. |
| [AOT and trimming](/docs/advanced/aot-trimming/) | Review runtime and deployment guidance for lean extraction services. |
| [OfficeIMO.Markdown](/products/markdown/) | Pair extraction with markdown rendering and transformation workflows. |
