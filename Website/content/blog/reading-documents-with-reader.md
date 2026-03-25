---
title: "Reading Any Office Document with OfficeIMO.Reader"
description: "Learn how to use OfficeIMO.Reader to extract text and metadata from Office documents for AI ingestion, search indexing, and batch processing."
date: 2026-03-01
tags: [reader, ingestion, ai]
categories: [Tutorial]
author: "Przemyslaw Klys"
---

Large language models and search engines are hungry for text, but the text they need is locked inside DOCX, XLSX, and PPTX files scattered across file shares and SharePoint libraries. **OfficeIMO.Reader** provides a unified API to crack open any supported Office format and extract structured text, ready for embedding, indexing, or summarisation.

## Installation

```bash
dotnet add package OfficeIMO.Reader
```

OfficeIMO.Reader has no native dependencies. It works on Windows, Linux, and macOS.

## Reading a Single Document

```csharp
using OfficeIMO.Reader;

var result = DocumentReader.Read("proposal.docx");

Console.WriteLine($"Format:    {result.Format}");       // Docx
Console.WriteLine($"Pages:     {result.PageCount}");
Console.WriteLine($"Words:     {result.WordCount}");
Console.WriteLine($"Text:\n{result.Text}");
```

The `Read` method detects the format from magic bytes, not the file extension, so renamed or extension-less files are handled correctly.

## Extracting Metadata

```csharp
var meta = result.Metadata;

Console.WriteLine($"Title:     {meta.Title}");
Console.WriteLine($"Author:    {meta.Author}");
Console.WriteLine($"Created:   {meta.Created}");
Console.WriteLine($"Modified:  {meta.Modified}");
Console.WriteLine($"Keywords:  {string.Join(", ", meta.Keywords)}");
```

Metadata is normalised across formats. Whether the source is DOCX, XLSX, or PPTX, you get the same `DocumentMetadata` type.

## Batch Extraction

Processing a folder of mixed documents is a common ingest scenario:

```csharp
using OfficeIMO.Reader;

var files = Directory.GetFiles(@"/data/documents", "*.*", SearchOption.AllDirectories)
    .Where(f => DocumentReader.IsSupported(f));

var results = new List<ExtractionResult>();

Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 8 }, file =>
{
    try
    {
        var result = DocumentReader.Read(file);
        lock (results) results.Add(new(file, result));
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Failed: {file} - {ex.Message}");
    }
});

Console.WriteLine($"Extracted text from {results.Count} documents.");

record ExtractionResult(string Path, ReadResult Result);
```

Parallel extraction keeps all cores busy. Each call to `DocumentReader.Read` is stateless, so there is no contention.

## Chunking for AI Pipelines

LLMs have token limits. You cannot feed a 200-page contract into a single prompt. OfficeIMO.Reader includes a chunking utility that splits extracted text into overlapping segments:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Chunking;

var result = DocumentReader.Read("contract.docx");

var chunks = TextChunker.Chunk(result.Text, new ChunkOptions
{
    MaxTokens = 512,
    OverlapTokens = 64,
    TokenizerModel = "cl100k_base"  // GPT-4 tokenizer
});

foreach (var chunk in chunks)
{
    Console.WriteLine($"Chunk {chunk.Index}: {chunk.TokenCount} tokens");
    // Send chunk.Text to embedding API
}
```

The `TokenizerModel` parameter selects the tokenizer used to count tokens. Supported models include `cl100k_base` (GPT-4) and `o200k_base` (GPT-4o).

## Token Estimation

If you just need a quick count without full chunking:

```csharp
int tokens = TokenEstimator.Estimate(result.Text, "cl100k_base");
Console.WriteLine($"Estimated tokens: {tokens}");
```

This is useful for filtering out documents that are too large before sending them to a rate-limited API.

## Building a Search Index

Combine extraction with a search library like Lucene.NET:

```csharp
foreach (var (path, result) in results)
{
    var doc = new Document();
    doc.Add(new StringField("path", path, Field.Store.YES));
    doc.Add(new TextField("content", result.Text, Field.Store.NO));
    doc.Add(new StringField("author", result.Metadata.Author ?? "", Field.Store.YES));
    writer.AddDocument(doc);
}
```

Every document in your file share becomes searchable in seconds.

## Supported Formats

| Format | Extension | Status |
|---|---|---|
| Word (Open XML) | .docx | Full support |
| Excel (Open XML) | .xlsx | Full support |
| PowerPoint (Open XML) | .pptx | Full support |
| Rich Text | .rtf | Text extraction |
| Plain Text | .txt | Pass-through |

Legacy binary formats (.doc, .xls, .ppt) are not supported. For those, consider a pre-conversion step using LibreOffice.

## Conclusion

OfficeIMO.Reader bridges the gap between unstructured Office files and structured text pipelines. Whether you are building RAG for an LLM, a full-text search index, or a compliance scanner, Reader gives you clean text and metadata with a single method call.
