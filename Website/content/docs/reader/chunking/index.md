---
title: Chunk Documents for Search and AI
description: "Split documents into bounded, citable chunks with heading paths and page locations for search indexes and retrieval with OfficeIMO.Reader."
order: 57
packages:
  - OfficeIMO.Reader.Core
cmdlets:
  - Read-OfficeDocument
limits: "Token counts from the built-in estimator are approximate; use your model's tokenizer when a hard budget matters."
search_terms:
  - chunk documents
  - rag
  - embeddings
  - ai ingestion
verified: 2026-11-14
---

Search indexes and language-model retrieval work on passages, not whole files. OfficeIMO.Reader produces two kinds of chunks:

- **Flat chunks** bounded by characters (`ReaderOptions.MaxChars`). Each chunk carries its location: heading path, page, slide, sheet and cell range, or line numbers.
- **Hierarchical chunks** bounded by tokens, with optional overlap and the heading context of each passage. Use these when chunks are embedded for a model with a token budget.

## Flat chunks with locations

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .AddPdfHandler()
    .Build();

ReaderOptions options = ReaderOptions.CreateSafeIngestion();
options.MaxChars = 2_000;

foreach (ReaderChunk chunk in reader.Enumerate("Handbook.pdf", options)) {
    string where = chunk.Location.Page is int page ? $"page {page}" : chunk.Location.HeadingPath ?? "";
    Console.WriteLine($"{chunk.Id} [{where}] ~{chunk.TokenEstimate} tokens");
    // index chunk.Text or chunk.Markdown together with chunk.Location and chunk.ChunkHash
}
```

`ChunkHash` changes only when the chunk text changes, so an indexer can skip unchanged passages on the next run.

### PowerShell

```powershell
Read-OfficeDocument -Path .\Handbook.pdf -As Chunks -MaximumChunkCharacters 2000 |
    Select-Object Id, Kind, TokenEstimate
```

## Token-sized hierarchical chunks with overlap

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

ReaderChunkHierarchyResult result = reader.ReadHierarchical("Handbook.docx",
    chunkingOptions: new ReaderHierarchicalChunkingOptions {
        MaxTokens = 500,
        OverlapTokens = 50,
        IncludeContextInText = true
    });

foreach (ReaderChunk chunk in result.Chunks) {
    Console.WriteLine($"{chunk.Location.HeadingPath}: {chunk.Text.Length} characters");
}
```

The built-in counter estimates one token per four characters. To size chunks with your model's tokenizer, pass `TokenCounter = new ReaderDelegateTokenCounter("my-model:vocab-1", text => tokenizer.CountTokens(text))`.

### PowerShell

```powershell
$options = New-OfficeDocumentHierarchyOptions -MaximumTokens 500 -OverlapTokens 50 -IncludeContextInText
Read-OfficeDocument -As Hierarchy -Path .\Handbook.pdf -HierarchyOptions $options
```

## Chunk a result you already read

If you already have an `OfficeDocumentReadResult`, for example from `ReadDocument` or from stored JSON, chunk it without reading the file again:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
OfficeDocumentReadResult document = reader.ReadDocument("Handbook.docx");
ReaderChunkHierarchyResult chunks = reader.ChunkDocument(document,
    new ReaderHierarchicalChunkingOptions { MaxTokens = 800 });
```
