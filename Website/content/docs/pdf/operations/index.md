---
title: "PDF Inspection, Extraction, and Operations"
description: "Preflight, repair, extract, merge, split, stamp, optimize, redact, and inspect PDFs with structured reports. Includes examples and API links."
layout: docs
---

`OfficeIMO.Pdf` can operate on existing PDFs as well as author new ones. The read path exposes preflight and diagnostic reports so an application can distinguish a readable file from one that required recovery or cannot be safely rewritten.

## Start with a check

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("incoming.pdf");
PdfAnalysisReport analysis = pdf.Analyze();
if (!analysis.CanRead) {
    throw new InvalidDataException("The PDF cannot be read safely.");
}
```

`Analyze()` combines health, diagnostics, repair, optimization and signature evidence. Treat `CanRead` and `CanRewrite` separately: extraction may be possible even when a safe mutation is not. `AssessMutations()` lists which page, form, metadata and signing operations are allowed for this file.

## Extract content and assets

```csharp
PdfDocumentReadResult result = PdfDocument.Load("Service-review.pdf").Read();
foreach (PdfLogicalPage page in result.Pages) {
    Console.WriteLine($"Page {page.PageNumber}: {page.Paragraphs.Count} paragraphs, {page.Tables.Count} tables");
}
Console.WriteLine(result.ExtractMarkdown());
```

`Read()` returns paragraphs, headings, lists, tables, links, form fields and attachment metadata per page. `pdf.Images.Extract()` and `pdf.Attachments.Extract()` return image and embedded-file bytes, and Reader adapters normalize PDF content with other document families for ingestion pipelines.

## Compose existing documents

```csharp
PdfDocument.Load("Service-review.pdf")
    .MergeWith("Appendix.pdf")
    .Pages.Rotate(90, "2")
    .Save("Review-pack.pdf");

IReadOnlyList<PdfDocument> parts = PdfDocument.Load("Review-pack.pdf").Pages.Split();
for (int i = 0; i < parts.Count; i++) parts[i].Save($"Review-pack-{i + 1}.pdf");
```

Page operations return a new document; the source file is never changed in place. `PdfDocument.Merge(documents, options, out PdfMergeReport report)` combines several loaded documents under an explicit structure policy, and `pdf.Pages.SplitForProduction(...)` splits by page count, boundary text or target size.

Stamp and watermark operations support text, images and imported pages. Annotation editors, page operations, metadata changes, and optimization produce mutation evidence where signatures or structural constraints matter.

## Redaction is removal, not decoration

Use `pdf.Redactions.Search(...)` to build a reviewable plan, then `ApplyWithEvidence(plan)` to remove the content and verify the saved bytes. Drawing a black rectangle over text is not redaction; the original content may remain extractable.

## Operational policy

Set input-size, page-count, decompression, attachment, and timeout limits for untrusted documents. Keep repair diagnostics, mutation reports, and verification results with the job record when reproducibility matters.
