---
title: "Search and RAG Across Documents"
description: "Create deterministic search and RAG chunks from Office, PDF, email, OneNote, EPUB, HTML, and other formats while preserving source locations."
meta.eyebrow: "Mixed-document search and RAG"
meta.outcome: "One inspectable extraction contract across installed formats"
meta.primary_label: "Read the Reader guide"
meta.primary_url: "/docs/reader/"
---

OfficeIMO.Reader normalizes documents into text, Markdown, tables, metadata, assets, visuals, forms, diagnostics, and source-aware chunks. Format parsing remains in the owning OfficeIMO package; Reader supplies the common read-only contract.

This separation matters for search and RAG. Extraction should be deterministic and inspectable before an embedding model, vector database, classifier, or language model interprets the content.

## Install the formats the service accepts

```powershell
dotnet add package OfficeIMO.Reader.Word
dotnet add package OfficeIMO.Reader.Pdf
dotnet add package OfficeIMO.Reader.Email
dotnet add package OfficeIMO.Reader.OneNote
```

Use `OfficeIMO.Reader.All` only when the host deliberately accepts the complete local adapter family. Selective packages reduce the dependency and deployment surface for a service with a fixed input contract.

## Preserve the path back to evidence

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.OneNote;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .AddPdfHandler()
    .AddEmailHandlers()
    .AddOneNoteHandler()
    .Build();

ReaderChunkHierarchyResult hierarchy = reader.ReadHierarchical(
    "policy.docx",
    chunkingOptions: new ReaderHierarchicalChunkingOptions {
        MaxTokens = 800,
        OverlapTokens = 80,
        MaxInputChunks = 10_000,
        MaxOutputChunks = 50_000
    });

foreach (ReaderChunk chunk in hierarchy.Chunks) {
    Console.WriteLine($"{chunk.Id}: {chunk.TokenEstimate ?? 0} tokens");
}
```

Chunks retain deterministic IDs and hashes, source spans, heading ancestry, page, slide, sheet, section, or format-specific locations when available. Store those fields with the embedding so a search result or generated answer can point back to its source.

## Keep the AI provider optional

Reader does not choose a model, vector store, or hosted service. The application decides whether extracted content stays local, is redacted, enters an on-premises index, or is sent to a third party. That boundary keeps document parsing reusable and makes privacy choices visible.

## Plan difficult inputs

- Use content detection when uploads may be mislabeled.
- Set byte, document, folder, table, output, and concurrency limits.
- Treat encrypted, malformed, unsupported, and oversized inputs as recorded outcomes.
- Keep native text and OCR as separate source layers.
- Supply the embedding model's tokenizer when exact token counts matter.
- Retain structured diagnostics beside the indexed record.

Continue with the [Reader documentation](/docs/reader/), [document ingestion overview](/solutions/document-ingestion/), or [Reader package catalog](/docs/capabilities/packages/).
