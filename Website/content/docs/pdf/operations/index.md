---
title: "PDF Inspection, Extraction, and Operations"
description: "Preflight, repair, extract, merge, split, stamp, optimize, convert to PDF/A, redact, and inspect PDFs with structured reports. Includes examples and API links."
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

Page operations return a new document and do not overwrite the source; save the result to a different path to keep the original. `PdfDocument.Merge(documents, options, out PdfMergeReport report)` combines several loaded documents under an explicit structure policy, and `pdf.Pages.SplitForProduction(...)` splits by page count, boundary text or target size.

Stamp and watermark operations support text, images and imported pages. Annotation editors, page operations, metadata changes, and optimization produce mutation evidence where signatures or structural constraints matter.

## Convert an existing PDF to PDF/A

```csharp
PdfAConversionResult archive = PdfDocument.Load("Service-review.pdf").ConvertToPdfA(new PdfAConversionOptions {
    Profile = PdfComplianceProfile.PdfA3B,
    Attachments = PdfAAttachmentPolicy.Refuse
});
if (!archive.Succeeded) {
    foreach (PdfAConversionViolation violation in archive.Report.Violations) {
        Console.WriteLine($"{violation.Reason} at {violation.Location}: {violation.Message}");
    }
} else {
    archive.Save("Service-review-archive.pdf");
}
```

Conversion targets PDF/A-2b, 2u, 3b, 3u, 4 or 4f. It fixes what it can without moving anything on the page and lists every change in `Report.Changes`. When something cannot be made compliant without changing the content, nothing is produced and `Report.Violations` lists every reason by location and object number. `SaveAsPdfA(path)` does the same in one call and throws `PdfAConversionException` on refusal.

The same conversion is available as `OfficeWorkflow.ConvertToPdfA` in `OfficeIMO.Workflows`, as `officeimo pdf convert-to-pdfa` in the command-line tool, and as the `officeimo_pdf_convert_to_pdfa` MCP tool.

## Print to an installed printer

`OfficeIMO.Workflows` prints through the operating system's own queues: GDI on Windows and CUPS on macOS and Linux. It plans and renders the sheets once, so a preview shows the same pixels the printer receives. Submission sends every prepared sheet, not only the ones you previewed.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

var printers = new PdfPrinterService();
PdfDocument document = await PdfDocument.LoadAsync("report.pdf");
PdfPreparedPrintDocument prepared = PdfPrintRenderer.Prepare(document, new PdfPrintPlanRequest {
    InputPath = "report.pdf",
    Pages = "1-8",
    PagesPerSheet = 2
});
File.WriteAllBytes("sheet-1.png", prepared.Sheets[0].GetPng());

PdfPrintSubmission job = await printers.SubmitAsync(prepared, new PdfPrintDeliveryOptions {
    PrinterName = "Office printer",
    Copies = 2,
    Duplex = PdfPrintDuplex.LongEdge
});
```

- **Layout:** 1, 2, 4, 6 or 9 pages per sheet, odd or even pages, fit, actual size, fill or a custom percentage, per-edge margins, alignment, and grayscale.
- **Printer:** list queues with `GetPrintersAsync` and their trays with `GetPaperSourcesAsync`, then choose copies, duplex and paper source.
- **Permissions and limits:** PDF print restrictions are enforced, and page, pixel and output budgets bound the work.
- **Receipt:** the job identifier confirms that the queue accepted the job, not that paper came out. After an interrupted submission, check the queue before retrying.

From the command line, use `officeimo workflow printers` and `officeimo workflow print report.pdf --printer "Office printer"`.

## Redaction is removal, not decoration

Use `pdf.Redactions.Search(...)` to build a reviewable plan, then `ApplyWithEvidence(plan)` to remove the content and verify the saved bytes. Call `ThrowIfUnverified()` on the result before saving or sharing it: evidence can report residual or inconclusive matches. Drawing a black rectangle over text is not redaction; the original content may remain extractable.

## Operational policy

Set input-size, page-count, decompression, attachment, and timeout limits for untrusted documents. Keep repair diagnostics, mutation reports, and verification results with the job record when reproducibility matters.
