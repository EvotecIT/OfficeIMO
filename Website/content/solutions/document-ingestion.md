---
title: "Extract Documents for Search, Review, and AI"
description: "Normalize text, tables, metadata, images, and chunks from Office, PDF, email, Markdown, CSV, and other formats with OfficeIMO.Reader."
meta.eyebrow: "Document ingestion"
meta.outcome: "One inspectable read model across document families"
meta.primary_label: "Start with OfficeIMO.Reader"
meta.primary_url: "/docs/reader/"
meta.powershell: true
---

Search and AI systems usually do not need a full editable document model. They need reliable text, structure, metadata, tables, images, and source locations. `OfficeIMO.Reader` provides a normalized read-only layer over the focused OfficeIMO engines so ingestion code does not have to invent a separate parser for each file family.

## Keep extraction and interpretation separate

Use the reader to produce stable content blocks and metadata. Apply indexing, classification, summarization, embeddings, review rules, or retention policy afterward. This separation makes it possible to inspect what came from the file before an AI model interprets it.

Reader adapters cover Word, Excel, PowerPoint, PDF, Markdown, CSV, HTML, email, OneNote, and additional modular formats. Applications can install only the adapters they need.

## Preserve source context

Useful ingestion retains where content came from: document metadata, page or section context, worksheet and cell ranges, slide identity, table structure, and other available source locations. Store those references with indexed chunks so search results and generated answers can point back to evidence.

## Treat difficult files deliberately

Encrypted, malformed, unsupported, or oversized files should produce a visible outcome rather than disappearing from a batch. Record parser diagnostics, apply size and resource limits, and decide which failures can be retried or require review.

Legacy DOC, XLS, XLSB, and PPT inputs are handled by the owning Office engines. Consult the [format compatibility evidence](/compatibility/) when an ingestion workflow depends on a specific feature rather than plain content extraction.

## Privacy stays architectural

OfficeIMO runs inside your process and does not require sending files to a hosted conversion service. If you later send extracted content to a search platform or AI provider, that is a separate choice your application can govern, log, redact, or keep entirely local.

See the [Reader product](/products/reader/), [Reader guide](/docs/reader/), and [package catalog](/downloads/) for adapter selection.
