---
title: Reader and Extraction
description: Overview of the OfficeIMO.Reader package for unified document extraction and AI-ready chunking workflows.
order: 55
---

# Reader and Extraction

`OfficeIMO.Reader` provides a single extraction surface for Office and adjacent document formats. Instead of maintaining separate parsing pipelines for `.docx`, `.xlsx`, `.pptx`, Markdown, or PDF files, you can normalize them into one structured result model and then feed that output into indexing, search, and AI workflows.

## Best fit scenarios

- Build ingestion pipelines for RAG, semantic search, or compliance review.
- Normalize mixed document folders into one extraction and chunking model.
- Preserve headings, citations, and metadata while preparing content for downstream tools.
- Run extraction in background workers, containers, Azure Functions, or scheduled jobs.

## Core workflow

1. Extract a file into a structured result with text, metadata, and source information.
2. Chunk the result into deterministic slices sized for search or AI prompts.
3. Store chunks, citations, and metadata in your vector store, search index, or audit trail.

## Quick start

```csharp
using OfficeIMO.Reader;

var extraction = DocumentReader.Extract("proposal.docx");

Console.WriteLine(extraction.Title);
Console.WriteLine(extraction.Text.Length);

var chunks = DocumentReader.Chunk("proposal.docx", new ChunkOptions
{
    MaxTokens = 512,
    Overlap = 64,
    PreserveHeadings = true
});

foreach (var chunk in chunks)
{
    Console.WriteLine($"{chunk.Index}: {chunk.Heading}");
    Console.WriteLine(chunk.Citation);
}
```

## Formats and behavior

| Input | Typical use |
|-------|-------------|
| Word (`.docx`) | Rich business documents, reports, contracts, and templates |
| Excel (`.xlsx`) | Workbook content, tabular reports, and structured exports |
| PowerPoint (`.pptx`) | Slide decks, speaker notes, and presentation narratives |
| Markdown | Documentation, changelogs, developer notes, and generated content |
| PDF | Published exports, archival documents, and third-party handoffs |

## Design goals

- **Deterministic chunking** so repeated runs produce stable chunk boundaries.
- **Heading-aware extraction** so downstream systems retain document structure.
- **Citation-friendly metadata** so search and AI responses can reference original sources.
- **Container-friendly execution** with no Office installation requirements.

## Related packages

- [OfficeIMO.Word](/products/word/) for producing `.docx` content before ingestion.
- [OfficeIMO.Markdown](/products/markdown/) for rendering, transforming, and re-emitting extracted content.
- [AOT and Trimming](/docs/advanced/aot-trimming/) for lean deployment guidance.
