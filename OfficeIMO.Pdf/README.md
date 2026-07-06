# OfficeIMO.Pdf - Dependency-free PDF engine

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf)](https://www.nuget.org/packages/OfficeIMO.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Pdf)

`OfficeIMO.Pdf` is the first-party PDF package for OfficeIMO. It creates, reads, inspects, edits, merges, splits, stamps, and exports PDFs without runtime package dependencies.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Pdf
```

## Quick start

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(new PdfOptions {
        DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = 11
    })
    .Meta(title: "Hello PDF", author: "OfficeIMO")
    .H1("OfficeIMO.Pdf")
    .Paragraph(p => p
        .Text("A dependency-free PDF builder with ")
        .Bold("rich text")
        .Text(", links, tables, images, and document operations."))
    .Table(new[] {
        new[] { "Area", "Status" },
        new[] { "Runtime dependencies", "None in OfficeIMO.Pdf" },
        new[] { "License", "MIT" }
    })
    .Save("hello.pdf");
```

## What it does

- Creates PDFs with page setup, headings, paragraphs, rich text, links, lists, panels, rows/columns, tables, images, vector drawing, headers, footers, watermarks, metadata, and form primitives.
- Reads and inspects PDFs through text extraction, logical document objects, page metadata, links, images, attachments, outlines, forms, active-content diagnostics, and security/revision markers.
- Manipulates existing PDFs with page extraction, split, merge, delete, duplicate, move, rotate, metadata editing, stamps, and watermarks while preserving source PDF header versions on shared rewrite paths.
- Provides conversion reports, grouped warning summaries, and diagnostics so adapters can expose unsupported or simplified source content honestly.
- Provides reusable conversion proof snapshots for generated PDFs, artifact hashes, required page counts, page sizes, document metadata, outline titles, URI links, form fields, named destinations, page labels, attachments, output intents, optional-content/layer metadata, catalog/viewer metadata, XMP/tagged metadata, text markers, logical readback signals, expected and accepted warning contracts, and post-processing hand-off.
- Provides reusable rewrite-preservation proof for page geometry, metadata, navigation, catalog/viewer/action state, optional content, tagged content, security signatures, document versions, and source-structure markers such as incremental updates, xref streams, and object streams.
- Provides a reusable rewrite-preservation matrix for classifying named manipulation scenarios as rewrite-safe, preservation-failed, blocked by safety checks, or operation-failed, including optional-content/layer drift, targeted form-fill preservation, form/tagged/active-content/signature blockers, and fluent `PdfDocument` helpers for normal document rewrite operations.
- Serves as the shared engine for Word, Excel, Markdown, HTML, and PowerPoint PDF adapters.

## Existing PDF workflows

```csharp
using OfficeIMO.Pdf;

PdfDocument.Open("input.pdf")
    .Pages.Extract("1-2,4")
    .MergeWith("appendix.pdf")
    .UpdateMetadata(title: "Merged report")
    .Stamp.Text("Reviewed")
    .Save("output.pdf");

string text = PdfDocument.Open("output.pdf").Read.Text();
```

## Examples

### Write a generated PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(new PdfOptions {
        PageSize = PageSizes.A4,
        Margins = PageMargins.UniformCentimeters(1.6),
        DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = 10
    })
    .Meta(
        title: "Service report",
        author: "OfficeIMO",
        subject: "Generated PDF")
    .Header(h => h.AlignCenter().Text("Service report"))
    .Footer(f => f.AlignRight().Text("Page {page} of {pages}"))
    .H1("Service report")
    .Paragraph(p => p
        .Text("Generated ")
        .Bold(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm 'UTC'"))
        .Text(" with first-party PDF primitives."))
    .Table(new[] {
        new[] { "System", "Status", "Owner" },
        new[] { "Identity", "Green", "Operations" },
        new[] { "Messaging", "Yellow", "Exchange" }
    })
    .Save("service-report.pdf");
```

### Rich report layout

```csharp
PdfDocument.Create()
    .H1("Operational summary")
    .Paragraph(p => p
        .Text("Generated ")
        .Bold(DateTime.Today.ToString("yyyy-MM-dd"))
        .Text(" with links, lists, panels, and tables."))
    .Bullets(list => list
        .Item("No runtime package dependencies")
        .Item("Word-like document flow")
        .Item("Reusable PDF primitives for adapters"))
    .Panel(panel => panel
        .H2("Review note")
        .Paragraph(p => p.Text("Keep polished report designs in samples; keep reusable primitives in the engine.")))
    .Table(new[] {
        new[] { "Area", "Status" },
        new[] { "Layout", "Ready" },
        new[] { "Reading", "Evolving" }
    })
    .Save("summary.pdf");
```

### Read text, Markdown, tables, images, and attachments

```csharp
using OfficeIMO.Pdf;

using var pdf = PdfDocument.Open("statement.pdf");

string text = pdf.Read.Text();
string firstPages = pdf.Read.Text("1-2");
PdfOperationResult<string> safeFirstPages = pdf.Read.TryText("1-2");
string markdown = pdf.Read.Markdown();
IReadOnlyList<string> pages = pdf.Read.TextByPage();
PdfLogicalDocument logical = pdf.Read.Logical();
PdfMetadata metadata = pdf.Read.Metadata();
PdfDocumentSecurityInfo security = pdf.Read.Security();
IReadOnlyList<PdfPageInfo> pageInfo = pdf.Read.Pages();
PdfXmpMetadataInfo? xmp = pdf.Read.XmpMetadata();
IReadOnlyList<PdfOutputIntentInfo> outputIntents = pdf.Read.OutputIntents();
PdfTaggedContentInfo? taggedContent = pdf.Read.TaggedContent();
PdfOptionalContentProperties? optionalContent = pdf.Read.OptionalContent();
PdfOperationResult<PdfDocumentInfo> safeInfo = pdf.Read.TryDocumentInfo();

foreach (var table in logical.Tables) {
    Console.WriteLine($"Table on page {table.PageNumber}: {table.Rows.Count} rows");
}

string markdownTables = PdfLogicalTableTextExportExtensions.ExtractMarkdownTables("statement.pdf");
IReadOnlyList<PdfExtractedImage> images = pdf.Read.Images();
IReadOnlyList<PdfExtractedImage> firstPageImages = pdf.Read.Images("1");
PdfOperationResult<IReadOnlyList<PdfExtractedImage>> safeImages = pdf.Read.TryImages("1-2");
IReadOnlyList<PdfImagePlacement> imageGeometry = pdf.Read.ImagePlacements("1-2");
IReadOnlyList<PdfOutlineItem> outlines = pdf.Read.Outlines();
IReadOnlyList<PdfLogicalLinkAnnotation> links = pdf.Read.Links();
IReadOnlyList<PdfLogicalLinkAnnotation> supportLinks = pdf.Read.LinksByUri("https://example.com/support");
PdfOperationResult<IReadOnlyList<PdfNamedDestination>> safeDestinations = pdf.Read.TryNamedDestinations();
IReadOnlyList<PdfAnnotation> annotations = pdf.Read.Annotations();
IReadOnlyList<PdfAnnotation> freeTextNotes = pdf.Read.AnnotationsBySubtype("FreeText");
PdfOperationResult<IReadOnlyList<PdfAnnotation>> safeAnnotations = pdf.Read.TryAnnotations();
IReadOnlyList<PdfCatalogAction> catalogActions = pdf.Read.CatalogActions();
IReadOnlyList<PdfPageAction> pageActions = pdf.Read.PageActions();
PdfOperationResult<IReadOnlyList<PdfPageAction>> safePageActions = pdf.Read.TryPageActions();
IReadOnlyList<PdfFormField> formFields = pdf.Read.FormFields();
IReadOnlyList<PdfLogicalFormWidget> formWidgets = pdf.Read.FormWidgets("Person.Name");
PdfOperationResult<IReadOnlyList<PdfFormField>> safeFormFields = pdf.Read.TryFormFields();
IReadOnlyList<PdfAttachmentInfo> attachmentMetadata = pdf.Read.AttachmentMetadata();
IReadOnlyList<PdfExtractedAttachment> attachments = pdf.Read.Attachments();
PdfOperationResult<IReadOnlyList<PdfExtractedAttachment>> safeAttachments = pdf.Read.TryAttachments();
```

### Split and extract pages

```csharp
using OfficeIMO.Pdf;

using var source = PdfDocument.Open("packet.pdf");

source.Pages.Extract("1-3")
    .Save("cover-and-summary.pdf");

IReadOnlyList<PdfDocument> singlePageDocuments = source.Pages.Split();
for (int index = 0; index < singlePageDocuments.Count; index++) {
    singlePageDocuments[index].Save($"packet-page-{index + 1:000}.pdf");
}

IReadOnlyList<PdfDocument> selectedRanges = source.Pages.Split("1-2,5-6");
selectedRanges[0].Save("packet-front.pdf");
selectedRanges[1].Save("packet-evidence.pdf");
```

### Merge, reorder, delete, duplicate, move, and rotate

```csharp
using OfficeIMO.Pdf;

PdfDocument.Open("packet.pdf")
    .MergeWith("appendix.pdf")
    .Pages.Delete("2,5-6")
    .Pages.Duplicate("1")
    .Pages.Move(insertBeforePageNumber: 3, pageRanges: "7-8")
    .Pages.Rotate(90, "4")
    .UpdateMetadata(title: "Cleaned packet")
    .Save("packet-clean.pdf");
```

### Stamp and watermark an existing PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Open("contract.pdf")
    .Stamp.Text("Reviewed", new PdfTextStampOptions {
        X = 72,
        Y = 720,
        FontSize = 18,
        Color = PdfColor.FromRgb(180, 30, 30)
    })
    .Stamp.TextWatermark("CONFIDENTIAL", new PdfTextStampOptions {
        FontSize = 54,
        Color = PdfColor.Gray,
        RotationDegrees = -35
    })
    .Save("contract-reviewed.pdf");
```

### Fill and flatten a PDF form

```csharp
using OfficeIMO.Pdf;

PdfDocument.Open("application-form.pdf")
    .Forms.FillAndFlatten(new Dictionary<string, string> {
        ["Applicant.Name"] = "Adele Vance",
        ["Applicant.Email"] = "adele@example.com",
        ["Approval.Status"] = "Approved"
    })
    .Save("application-form-filled.pdf");
```

### Assess compliance proof without overclaiming

```csharp
using OfficeIMO.Pdf;

PdfDocument document = PdfDocument.Create(new PdfOptions()
        .UsePdfA(PdfComplianceProfile.PdfA3B))
    .Paragraph(paragraph => paragraph.Text("Groundwork can be assessed before a formal claim."));

PdfComplianceProofReport proof = document.AssessComplianceProof(
    PdfComplianceProfile.PdfA3B,
    externalValidations: null);

if (!proof.CanClaimConformance) {
    Console.WriteLine(proof.ProofStatus);
    Console.WriteLine(proof.ExternalProofSummary);
}
```

### Choose converter-friendly text fallbacks

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("proposal.docx");

var options = new PdfSaveOptions {
    TextFallbacks = PdfTextFallbackFeatures.Default,
    AllowSystemFontEmbedding = true
}.UseProfile(PdfExportProfile.PrintReady);

var result = document.ToPdfDocumentResult(options);
result.ConversionReport.RequireNoErrorWarnings();
result.Save("proposal.pdf");
```

The Markdown, Word, Excel, and PowerPoint PDF adapters expose the same `TextFallbacks` enum. Use `PdfTextFallbackFeatures.None` when strict standard-font output is preferred, or `AllowSystemFontEmbedding = true` when the converter may embed installed host fonts for Unicode, symbols, and emoji.

### Add e-invoice groundwork

```csharp
using OfficeIMO.Pdf;

byte[] invoiceXml = File.ReadAllBytes("factur-x.xml");

PdfDocument.Create(new PdfOptions()
        .UseFacturX(invoiceXml, textFallbacks: PdfTextFallbackFeatures.DocumentFont))
    .Paragraph("Invoice preview")
    .Save("invoice.pdf");
```

### Page setup, watermarks, and metadata

```csharp
PdfDocument.Create(new PdfOptions {
        PageSize = PageSize.FromCentimeters(21, 29.7).Portrait(),
        Margins = PageMargins.UniformCentimeters(1.5),
        TextWatermark = new PdfTextWatermark("DRAFT") {
            Opacity = 0.12,
            RotationAngle = -35
        }
    })
    .Meta(title: "Draft report", author: "OfficeIMO")
    .H1("Draft report")
    .Paragraph("This document uses page-level options instead of post-processing.")
    .Save("draft.pdf");
```

### Inspect and preflight before rewriting

```csharp
using OfficeIMO.Pdf;

byte[] bytes = File.ReadAllBytes("incoming.pdf");
PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);

if (!preflight.Can(PdfPreflightCapability.ManipulatePages)) {
    foreach (string diagnostic in preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages)) {
        Console.WriteLine(diagnostic);
    }
}

var result = PdfDocument.Open(bytes).Pages.TryExtract("1-2");
if (result.Succeeded) {
    result.RequireValue().Save("incoming-first-pages.pdf");
}
```

### Inspect before automating

```csharp
using var pdf = PdfDocument.Open("incoming.pdf");

var inspection = pdf.Inspect();
Console.WriteLine($"Pages: {inspection.PageCount}");
Console.WriteLine($"Links: {inspection.LinkAnnotationCount}");
Console.WriteLine($"Forms: {inspection.FormFields.Count}");
Console.WriteLine($"Active content: {inspection.HasActiveContent}");

foreach (var page in inspection.Pages) {
    Console.WriteLine($"{page.PageNumber}: {page.Width} x {page.Height}");
}
```

### Convert PDFs through adapter packages

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var word = WordDocument.Load("proposal.docx");
word.SaveAsPdf("proposal.pdf");

PdfExcelTableConverterExtensions.SavePdfTablesAsExcel(
    "bank-statement.pdf",
    "bank-statement-tables.xlsx");

PdfHtmlConverter.SaveAsHtml(
    "proposal.pdf",
    "proposal-review.html",
    new PdfHtmlSaveOptions {
        Profile = PdfHtmlProfile.PositionedReview,
        IncludeLinkAnnotations = true,
        IncludeFormWidgets = true
    });
```

## Conversion adapters

| Package | Role |
| --- | --- |
| [OfficeIMO.Word.Pdf](../OfficeIMO.Word.Pdf/README.md) | Maps Word documents into PDF primitives. |
| [OfficeIMO.Excel.Pdf](../OfficeIMO.Excel.Pdf/README.md) | Maps Excel workbooks into PDF primitives. |
| [OfficeIMO.Markdown.Pdf](../OfficeIMO.Markdown.Pdf/README.md) | Maps Markdown documents into PDF primitives. |
| [OfficeIMO.PowerPoint.Pdf](../OfficeIMO.PowerPoint.Pdf/README.md) | Maps PowerPoint slides into PDF primitives. |
| [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) | Bridges HTML to PDF and PDF to HTML. |

## Boundaries

- `OfficeIMO.Pdf` should stay dependency-free at runtime. Rasterizers, visual comparison tools, and external renderers belong in tests or development tooling.
- Polished invoice, report, and statement examples belong in samples and visual fixtures, not as special engine concepts.
- Adapter-specific mapping belongs in the source adapter packages. Shared PDF layout, reading, and manipulation behavior belongs here.
- Current-state inventories belong in [Docs/officeimo.pdf.current-state.md](../Docs/officeimo.pdf.current-state.md), not in this NuGet README.

## Current state

The PDF engine is useful and broad, but it is still evolving. It has strong first-party coverage for common generated business documents and conservative read/manipulation workflows, while advanced typography, complex PDF preservation, encryption/decryption, and signature validation remain deeper roadmap areas.

For the full capability inventory and roadmap, read [Docs/officeimo.pdf.current-state.md](../Docs/officeimo.pdf.current-state.md).

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
