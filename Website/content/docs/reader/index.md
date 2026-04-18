---
title: Reader and Extraction
description: Overview of the OfficeIMO.Reader package for unified document extraction and AI-ready chunking workflows.
order: 55
---

# Reader and Extraction

`OfficeIMO.Reader` provides a single extraction surface for the built-in formats currently handled in the repo, plus optional adapters for adjacent document types. Instead of maintaining separate parsing pipelines for `.docx`, `.xlsx`, `.pptx`, Markdown, PDF, CSV, JSON, XML, HTML, EPUB, ZIP, or text-like files, you can normalize them into one chunk model and then feed that output into indexing, search, and AI workflows.

## Best fit scenarios

- Build ingestion pipelines for RAG, semantic search, or compliance review.
- Normalize mixed document folders into one extraction and chunking model.
- Preserve headings, citations, token estimates, and source hashes while preparing content for downstream tools.
- Run extraction in background workers, containers, Azure Functions, or scheduled jobs.

## Core workflow

1. Extract a file into `ReaderChunk` instances with text, markdown, tables, visuals, and source information.
2. Tune `ReaderOptions` so emitted slices stay deterministic and sized for search or AI prompts.
3. Store chunks, citations, and source identifiers in your vector store, search index, or audit trail.

## Quick start

```csharp
using OfficeIMO.Reader;

var chunks = DocumentReader.Read("proposal.docx", new ReaderOptions
{
    MaxChars = 4_000,
    IncludeWordFootnotes = true,
    ComputeHashes = true
}).ToList();

foreach (var chunk in chunks)
{
    Console.WriteLine($"{chunk.Id} :: {chunk.Kind}");
    Console.WriteLine(chunk.Location.HeadingPath ?? chunk.Location.Path);
    Console.WriteLine(chunk.TokenEstimate);
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
| CSV/TSV | Structured flat files through `OfficeIMO.Reader.Csv` |
| JSON | API exports and structured payloads through `OfficeIMO.Reader.Json` |
| XML | Element-oriented documents through `OfficeIMO.Reader.Xml` |
| HTML | Web pages and exported fragments through `OfficeIMO.Reader.Html` |
| EPUB | Books and packaged publications through `OfficeIMO.Reader.Epub` |
| ZIP | Archive traversal through `OfficeIMO.Reader.Zip` |

## Adapter packages

The core package covers the main Office-oriented extraction flow. Add adapter packages only when the deployment needs those input families:

- `OfficeIMO.Reader.Csv` for CSV and TSV ingestion.
- `OfficeIMO.Reader.Json` for JSON payloads.
- `OfficeIMO.Reader.Xml` for XML documents.
- `OfficeIMO.Reader.Text` when CSV, JSON, and XML adapters should travel together.
- `OfficeIMO.Reader.Html` for HTML normalization through the Markdown pipeline.
- `OfficeIMO.Reader.Epub` for EPUB publications.
- `OfficeIMO.Reader.Zip` for safe archive traversal.

Use the [Reader API reference](/api/reader/) for the core `DocumentReader`, `ReaderOptions`, chunk, manifest, and handler registration types. Adapter-specific APIs are documented through their package README files until the website generates separate adapter reference sections.

## Design goals

- **Deterministic chunking** so repeated runs produce stable chunk boundaries.
- **Heading-aware extraction** so downstream systems retain document structure.
- **Citation-friendly location data** so search and AI responses can reference original sources.
- **Incremental indexing support** through source IDs, hashes, and per-document chunk summaries.
- **Container-friendly execution** with no Office installation requirements.

## Related packages

- [OfficeIMO.Word](/products/word/) for producing `.docx` content before ingestion.
- [OfficeIMO.Markdown](/products/markdown/) for rendering, transforming, and re-emitting extracted content.
- [AOT and Trimming](/docs/advanced/aot-trimming/) for lean deployment guidance.
