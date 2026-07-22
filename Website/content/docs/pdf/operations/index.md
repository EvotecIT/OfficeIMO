---
title: "PDF Inspection, Extraction, and Operations"
description: "Preflight, repair, extract, merge, split, stamp, optimize, redact, and inspect PDFs with structured reports."
layout: docs
---

`OfficeIMO.Pdf` can operate on existing PDFs as well as author new ones. The read path exposes preflight and diagnostic reports so an application can distinguish a readable file from one that required recovery or cannot be safely rewritten.

## Start with preflight

```csharp
using OfficeIMO.Pdf;

PdfDocumentPreflight preflight = PdfDocument.Preflight("incoming.pdf");
if (!preflight.CanRead) {
    throw new InvalidDataException("The PDF cannot be read safely.");
}
```

Use `Diagnostics` when you need fonts, streams, structure, repair, security, and optimization evidence. Treat `CanRead` and `CanRewrite` separately: extraction may be possible even when a safe mutation is not.

## Extract content and assets

- `PdfTextExtractor` returns text by document, page, or page range and can project logical Markdown.
- `PdfImageExtractor` returns image bytes and placement metadata.
- `PdfAttachmentExtractor` enumerates and extracts embedded files.
- Reader adapters normalize PDF content with other document families for ingestion pipelines.

## Compose existing documents

`PdfMerger` and `PdfDocument.MergeWithReport` combine PDFs with an explicit structure policy and return readback evidence. `PdfPageExtractor` extracts individual pages or ranges and provides split helpers for paths, streams, and byte arrays.

Stamp and watermark operations support text and images. Annotation editors, page operations, metadata changes, and optimization produce mutation evidence where signatures or structural constraints matter.

## Redaction is removal, not decoration

Use the redaction planner and applier to remove targeted content, then run `PdfRedactionVerification` against the resulting bytes. Drawing a black rectangle over text is not redaction; the original content may remain extractable.

## Operational policy

Set input-size, page-count, decompression, attachment, and timeout limits for untrusted documents. Keep repair diagnostics, mutation reports, and verification results with the job record when reproducibility matters.
