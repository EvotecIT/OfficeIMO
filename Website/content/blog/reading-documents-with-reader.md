---
title: "Reading Any Office Document with OfficeIMO.Reader"
description: "Learn how to use OfficeIMO.Reader to extract normalized chunks from Office documents for AI ingestion, search indexing, and batch processing."
date: 2026-03-01
tags: [reader, ingestion, ai]
categories: [Tutorial]
author: "Przemyslaw Klys"
---

Large language models and search engines are hungry for text, but the text they need is locked inside DOCX, XLSX, PPTX, Markdown, and PDF files scattered across file shares and SharePoint libraries. **OfficeIMO.Reader** provides a unified API to crack open supported formats and emit normalized chunks, ready for embedding, indexing, or summarisation.

## Installation

```bash
dotnet add package OfficeIMO.Reader
```

OfficeIMO.Reader is designed for the same cross-platform .NET environments as the packages it builds on. It fits well in local tools, hosted services, and containerized ingestion jobs, but you should still validate the exact runtime and document mix you plan to ship.

## Reading a Single Document

```csharp
using OfficeIMO.Reader;

var chunks = DocumentReader.Read("proposal.docx", new ReaderOptions
{
    MaxChars = 4_000,
    IncludeWordFootnotes = true,
    ComputeHashes = true
}).ToList();

Console.WriteLine($"Chunks:    {chunks.Count}");

foreach (var chunk in chunks)
{
    Console.WriteLine($"Id:        {chunk.Id}");
    Console.WriteLine($"Kind:      {chunk.Kind}");
    Console.WriteLine($"Heading:   {chunk.Location.HeadingPath ?? "(root)"}");
    Console.WriteLine($"Tokens:    {chunk.TokenEstimate}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

For file paths, `DocumentReader.Read(...)` routes by extension and returns `ReaderChunk` items rather than one giant monolithic string. That makes it easier to preserve headings, pages, sheets, slides, and other citation-friendly location data.

## Inspecting Location And Structured Output

```csharp
var first = chunks.First();

Console.WriteLine($"Path:      {first.Location.Path}");
Console.WriteLine($"Heading:   {first.Location.HeadingPath}");
Console.WriteLine($"Page:      {first.Location.Page}");
Console.WriteLine($"Slide:     {first.Location.Slide}");
Console.WriteLine($"SourceId:  {first.SourceId}");
Console.WriteLine($"ChunkHash: {first.ChunkHash}");

if (first.Tables?.Count > 0)
{
    Console.WriteLine($"Tables:    {first.Tables.Count}");
}
```

Each chunk carries source and location data. Depending on the input kind, that can include page numbers, slide numbers, heading paths, A1 ranges, tables, visuals, warnings, and deterministic source identifiers for incremental indexing.

## Batch Extraction

Processing a folder of mixed documents is a common ingest scenario:

```csharp
using OfficeIMO.Reader;

var documents = DocumentReader.ReadFolderDocuments(
    folderPath: "/data/documents",
    folderOptions: new ReaderFolderOptions
    {
        Recurse = true,
        DeterministicOrder = true,
        MaxFiles = 1_000
    },
    options: new ReaderOptions
    {
        MaxChars = 4_000,
        ComputeHashes = true
    },
    onProgress: progress =>
        Console.WriteLine($"{progress.Kind}: scanned={progress.FilesScanned}, parsed={progress.FilesParsed}, skipped={progress.FilesSkipped}, chunks={progress.ChunksProduced}")
).ToList();

Console.WriteLine($"Visited:   {documents.Count} files");
Console.WriteLine($"Parsed:    {documents.Count(d => d.Parsed)}");
Console.WriteLine($"Chunks:    {documents.Sum(d => d.ChunksProduced)}");
```

`ReadFolderDocuments(...)` is the easiest handoff point for indexing pipelines because it gives you one `ReaderSourceDocument` per file, with `Chunks`, `SourceId`, `SourceHash`, warnings, and token totals already grouped.

## Controlling Chunk Size For AI Pipelines

LLMs have token limits. You cannot feed a 200-page contract into a single prompt. With OfficeIMO.Reader, you shape the emitted chunks through `ReaderOptions` instead of running a second tokenizer-specific chunker after extraction:

```csharp
using OfficeIMO.Reader;

var chunks = DocumentReader.Read("contract.docx", new ReaderOptions
{
    MaxChars = 3_000,
    IncludeWordFootnotes = false,
    MarkdownChunkByHeadings = true
}).ToList();

foreach (var chunk in chunks)
{
    Console.WriteLine($"{chunk.Id} :: {chunk.TokenEstimate} tokens");
    Console.WriteLine(chunk.Location.HeadingPath ?? chunk.Location.Path);
}
```

For Word and Markdown inputs, this keeps heading-aware chunk boundaries. For Excel and PowerPoint, the same options help keep sheet-, range-, and slide-oriented chunks small enough for downstream search or embedding jobs.

## Getting A Detailed Summary

If you want folder-level counts, skip reasons, and optional chunk materialization in one object:

```csharp
using OfficeIMO.Reader;

var result = DocumentReader.ReadFolderDetailed(
    folderPath: "/data/documents",
    folderOptions: new ReaderFolderOptions { Recurse = true, MaxFiles = 2_000 },
    options: new ReaderOptions { ComputeHashes = true, MaxChars = 4_000 },
    includeChunks: true,
    onProgress: progress => Console.WriteLine($"{progress.Kind}: {progress.Path}")
);

Console.WriteLine($"Files parsed: {result.FilesParsed}");
Console.WriteLine($"Files skipped: {result.FilesSkipped}");
Console.WriteLine($"Chunks produced: {result.ChunksProduced}");
Console.WriteLine($"Warnings: {result.Warnings?.Count ?? 0}");
```

This is useful for dashboards, scheduled jobs, and CI pipelines where you want more than raw chunks.

## Building a Search Index

Combine extraction with a search library like Lucene.NET:

```csharp
foreach (var source in documents.Where(d => d.Parsed))
{
    foreach (var chunk in source.Chunks)
    {
        var doc = new Document();
        doc.Add(new StringField("path", source.Path, Field.Store.YES));
        doc.Add(new StringField("sourceId", source.SourceId ?? "", Field.Store.YES));
        doc.Add(new TextField("content", chunk.Markdown ?? chunk.Text, Field.Store.NO));
        doc.Add(new StringField("heading", chunk.Location.HeadingPath ?? "", Field.Store.YES));
        writer.AddDocument(doc);
    }
}
```

Every document in your file share becomes searchable in chunk-sized slices, which usually works better for semantic search and citation-rich AI responses than indexing one giant text blob per file.

## Supported Formats

| Format | Extension | Status |
|---|---|---|
| Word (Open XML) | `.docx`, `.docm` | Built-in |
| Excel (Open XML) | `.xlsx`, `.xlsm` | Built-in |
| PowerPoint (Open XML) | `.pptx`, `.pptm` | Built-in |
| Markdown | `.md`, `.markdown` | Built-in |
| PDF | `.pdf` | Built-in |
| Text-like inputs | `.txt`, `.log`, `.csv`, `.tsv`, `.json`, `.xml`, `.yml`, `.yaml` | Built-in text reader |

Legacy binary formats (`.doc`, `.xls`, `.ppt`) are not supported. Structured adapters for formats like EPUB, HTML, ZIP, JSON, CSV, and XML can also be registered from optional `OfficeIMO.Reader.*` packages.

## Conclusion

OfficeIMO.Reader bridges the gap between unstructured Office files and indexing-friendly text pipelines. Whether you are building RAG for an LLM, a full-text search index, or a compliance scanner, Reader gives you normalized chunks, source identifiers, and citation-friendly locations with one API surface.
