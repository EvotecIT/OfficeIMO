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

- Creates PDFs with page setup, headings, paragraphs, rich text, links, lists, mixed inline images and boxes, dictionary-driven hyphenation, styled multipage containers, balanced block-flow columns, conditional/replayable flow, position capture, sections, generated TOCs, optional-content layers, tables, images, vector drawing, headers, footers, watermarks, metadata, portfolios, and form primitives.
- Reads and inspects PDFs through text extraction, logical document objects, page metadata, links, images, attachments, portfolios, outlines, forms, bounded immutable raw-structure views, active-content diagnostics, and security/revision markers.
- Manipulates existing PDFs with page extraction, split, merge, delete, duplicate, move, rotate, metadata editing, stamps, watermarks, and complete-page overlay/underlay while preserving source PDF header versions on shared rewrite paths.
- Renders supported embedded TrueType and OpenType/CFF fonts with stable-glyph subsetting. Built-in shaping remains dependency-free, while `IPdfTextShapingProvider` can supply positioned glyph advances and offsets for scripts that need a host-owned shaping engine.
- Shares managed CMYK, Lab, XYZ, calibrated-color conversion, vector tiling fills, standard blend modes, and alpha/luminosity soft masks with `OfficeIMO.Drawing`.
- Bounds completed page/effect content and serialized-object retention with separate memory limits, temporary-file spillover, direct large-stream spooling, and chunked final assembly during stream saves. Per-page metadata and the authored block model remain proportional to document size, and `ToBytes()` buffers the final artifact.
- Provides conversion reports, grouped warning summaries, and diagnostics so adapters can expose unsupported or simplified source content honestly.
- Provides reusable conversion proof snapshots for generated PDFs, artifact hashes, required page counts, page sizes, document metadata, outline titles, URI links, form fields, named destinations, page labels, attachments, output intents, optional-content/layer metadata, catalog/viewer metadata, XMP/tagged metadata, text markers, logical readback signals, expected and accepted warning contracts, and post-processing hand-off. Compliance proof records bind external validator name, version, profile, result, warnings, SHA-256, byte length, and validation time to the exact artifact.
- Provides reusable rewrite-preservation proof for page geometry, metadata, navigation, catalog/viewer/action state, optional content, tagged content, security signatures, document versions, and source-structure markers such as incremental updates, xref streams, and object streams.
- Provides a reusable rewrite-preservation matrix for classifying named manipulation scenarios as rewrite-safe, preservation-failed, blocked by safety checks, or operation-failed, including optional-content/layer drift, targeted form-fill preservation, form/tagged/active-content/signature blockers, and fluent `PdfDocument` helpers for normal document rewrite operations.
- Serves as the shared engine for Word, Excel, PowerPoint, Markdown, HTML, RTF, OneNote, AsciiDoc, and LaTeX PDF adapters.

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

`Open(...)` is the one entry point for byte arrays, files, and streams. It
enforces the same `PdfReadOptions` limits before buffering, snapshots caller
input once, and reuses one parsed document across read, inspection, preflight,
diagnostic, optimization, signature, and compliance operations.

For a single health and capability view:

```csharp
PdfAnalysisReport analysis = PdfDocument
    .Open("incoming.pdf")
    .Analyze(PdfComplianceProfile.PdfA2B);

Console.WriteLine($"Pages: {analysis.Info.PageCount}");
Console.WriteLine($"Readable: {analysis.CanRead}");
Console.WriteLine($"Rewrite safe: {analysis.CanRewrite}");
Console.WriteLine($"Healthy: {analysis.IsHealthy}");

foreach (PdfDiagnosticFinding finding in analysis.Diagnostics.Findings) {
    Console.WriteLine($"{finding.Severity}: {finding.Code} — {finding.Message}");
}
```

## Migrating to 3.0

Version 3.0 intentionally narrows the public API around the fluent `PdfDocument`
facade:

- Replace `PdfDocument.Load(...)` and `PdfReadDocument.Load(...)` with
  `PdfDocument.Open(...)` or `PdfReadDocument.Open(...)`.
- Seekable PDF input streams are now consistently read from the beginning and
  restored to their original position. Non-seekable streams are read forward
  from their current position.
- Keep one opened `PdfDocument` and reuse it for `Read`, `Inspect`, `Preflight`,
  `Analyze`, compliance, and manipulation work. The source snapshot and canonical
  parse are cached for that document.
- Use `PdfDocument.Analyze(...)` when a workflow needs the combined health,
  rewrite-safety, diagnostics, optimization, signature, repair, and compliance
  view.
- Use `CreateComplianceArtifact(...)` instead of separately rendering bytes and
  passing them back to `AssessComplianceProof(...)`. The returned immutable
  snapshot keeps exact output bytes and matching readiness evidence together,
  including for randomized encrypted output.
- Use the fluent `Pages`, `Forms`, `Attachments`, `Bookmarks`, `Annotations`,
  `Stamp`, `Security`, and metadata operations instead of the former public
  static engine classes. Those implementation engines are now internal so there
  is one supported route for each operation.

The package remains source-compatible across `netstandard2.0`, `net8.0`, and
`net10.0`; the API cleanup itself is a deliberate source-breaking change.

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

### Hyphenation and inline visuals

```csharp
byte[] statusIcon = File.ReadAllBytes("status.png");
var hyphenation = new PdfHyphenationLexicon(new[] {
    "auto-ma-tion",
    "ty-pog-ra-phy",
    "re-port-ing"
});

PdfDocument.Create(new PdfOptions()
        .UseTextHyphenationDictionary(hyphenation))
    .Paragraph(paragraph => paragraph
        .Text("Automation status ")
        .InlineImage(statusIcon, 12, 12, alternativeText: "Healthy")
        .Text(" remains available during long reporting runs."))
    .Save("inline-status.pdf");
```

Inline elements participate in normal line wrapping. In tagged output, image and box alternative text is carried into the structure tree.

### Sections, generated navigation, and bounded stream output

```csharp
var options = new PdfOptions {
    PageContentMemoryLimitBytes = 4 * 1024 * 1024,
    ObjectBufferMemoryLimitBytes = 8 * 1024 * 1024
};

PdfDocument.Create(options)
    .TableOfContents()
    .Section("Summary", section => section
        .Container(content => content
            .Paragraph(p => p.Text("A styled, keep-together summary."))))
    .Section("Details", section => section
        .Columns(columns => {
            columns.Paragraph(p => p.Text("First column"));
            columns.ColumnBreak();
            columns.Paragraph(p => p.Text("Second column"));
        }, new PdfMultiColumnOptions { ColumnCount = 2, Gap = 18 }))
    .Save("navigable-report.pdf");
```

### Read text, Markdown, tables, images, and attachments

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Open("statement.pdf");

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

PdfDocument source = PdfDocument.Open("packet.pdf");

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

Import a complete source page above or below selected target pages without
rasterizing it:

```csharp
PdfDocument.Open("contract.pdf")
    .Stamp.OverlayPage("letterhead.pdf", new PdfPageOverlayOptions {
        SourcePageNumber = 1,
        TargetPages = PdfPageSelector.Parse("all,!last"),
        Fit = PdfPageOverlayFit.Contain,
        Opacity = 0.9
    })
    .Save("contract-with-letterhead.pdf");
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

### Generate and assess validator-backed PDF/A

```csharp
using OfficeIMO.Pdf;

byte[] fontBytes = File.ReadAllBytes("SourceSerif4-Regular.otf");
var options = new PdfOptions()
    .UsePdfA(PdfComplianceProfile.PdfA2B)
    .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "Source Serif 4")
    .RequireCompliance(PdfComplianceProfile.PdfA2B);

PdfComplianceArtifact artifact = PdfDocument.Create(options)
    .Meta(title: "Archive copy")
    .Paragraph(paragraph => paragraph.Text("This artifact is ready for external validation."))
    .CreateComplianceArtifact(PdfComplianceProfile.PdfA2B);

byte[] pdf = artifact.ToBytes();
File.WriteAllBytes("archive.pdf", pdf);

// Create this result from the validator invocation in your build or release lane.
PdfExternalValidationResult validation = PdfExternalValidationResult.PassedForArtifact(
    PdfExternalValidatorKind.VeraPdf,
    "veraPDF",
    "1.30.2",
    "PDF/A-2b validation passed.",
    pdf,
    "PDF/A-2b");

PdfComplianceProofReport proof = artifact.AssessProof(new[] { validation });

if (!proof.CanClaimConformance) {
    throw new InvalidOperationException(proof.ExternalProofSummary);
}
```

Formal generation gates are available for PDF/A-2b, PDF/A-3b, PDF/UA-1, Factur-X, and ZUGFeRD. `RequireCompliance(...)` rejects incomplete generation settings. A conformance claim still requires a passing external result for the same profile, SHA-256, and byte length; validators are build-time tools and are not runtime dependencies of `OfficeIMO.Pdf`.

### Choose converter-friendly text fallbacks

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("proposal.docx");

var options = new PdfSaveOptions {
    TextFallbacks = PdfTextFallbackFeatures.Default,
    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost()
}.UseProfile(PdfExportProfile.PrintReady);

var result = document.ToPdfDocumentResult(options);
result.Report.RequireNoErrorWarnings();
result.Save("proposal.pdf");
```

The Word, Excel, PowerPoint, Markdown, HTML, RTF, OneNote, AsciiDoc, and LaTeX PDF adapters use one `PdfResourcePolicy`; semantic-projection adapters expose it through their nested Markdown PDF options. The balanced default enables installed fonts and bounded data URI/package resources for document fidelity while denying arbitrary local files and remote resolver calls. Use `PdfResourcePolicy.CreatePortableDeterministic()` for reproducible or untrusted conversion, and `CreateTrustedHost()` only when both source and host are trusted. Profiles never grant resource access.

The text-capable adapters also expose `TextFallbacks`. `PdfTextFallbackFeatures.Default` enables document, monospace, symbol, and emoji groups. Add `PdfTextFallbackFeatures.MultilingualFonts` for CJK, Arabic, and other non-Latin family candidates; OneNote adds that candidate group unless fallbacks are `None`. Candidate selection does not read installed fonts unless the resource policy allows it.

### Generate a formal e-invoice carrier

```csharp
using OfficeIMO.Pdf;

byte[] invoiceXml = File.ReadAllBytes("factur-x.xml");
byte[] fontBytes = File.ReadAllBytes("SourceSerif4-Regular.otf");

PdfDocument.Create(new PdfOptions()
        .UseFacturX(
            invoiceXml,
            relationship: PdfAssociatedFileRelationship.Alternative,
            textFallbacks: PdfTextFallbackFeatures.None)
        .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "Source Serif 4")
        .RequireCompliance(PdfComplianceProfile.FacturX))
    .Paragraph("Invoice preview")
    .Save("invoice.pdf");
```

The XML must be a valid EN 16931 CrossIndustryInvoice payload. The formal carrier gate checks the PDF/A-3 attachment, metadata, font, Unicode, and invoice rules before writing; exact-artifact PDF/A and invoice-validator results are still required before claiming conformance.

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
PdfDocument pdf = PdfDocument.Open(bytes);
PdfDocumentPreflight preflight = pdf.Preflight();

if (!preflight.Can(PdfPreflightCapability.ManipulatePages)) {
    foreach (string diagnostic in preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages)) {
        Console.WriteLine(diagnostic);
    }
}

var result = pdf.Pages.TryExtract("1-2");
if (result.Succeeded) {
    result.RequireValue().Save("incoming-first-pages.pdf");
}
```

### Inspect before automating

```csharp
PdfDocument pdf = PdfDocument.Open("incoming.pdf");

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

PdfExcelTableConverterExtensions.SaveAsExcel(
    "bank-statement.pdf",
    "bank-statement-tables.xlsx");

PdfHtmlConverterExtensions.SaveAsHtml(
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
| [OfficeIMO.Rtf.Pdf](../OfficeIMO.Rtf.Pdf/README.md) | Maps semantic RTF into PDF and logical PDF content back to RTF. |
| [OfficeIMO.OneNote.Pdf](../OfficeIMO.OneNote.Pdf/README.md) | Explicitly projects offline OneNote hierarchy into a semantic PDF document with loss diagnostics. |
| [OfficeIMO.AsciiDoc.Pdf](../OfficeIMO.AsciiDoc.Pdf/README.md) | Projects native AsciiDoc through the loss-aware Markdown bridge and combines parser, projection, and PDF diagnostics. |
| [OfficeIMO.Latex.Pdf](../OfficeIMO.Latex.Pdf/README.md) | Projects the bounded LaTeX profile through the loss-aware Markdown bridge without executing TeX. |

The canonical catalog records OpenDocument as manual loss-aware composition and formats that are intentionally not advertised as direct conversion yet: email, EPUB, and Visio. See [`Docs/pdf-conversion-scenarios.json`](../Docs/pdf-conversion-scenarios.json).

## Boundaries

- `OfficeIMO.Pdf` should stay dependency-free at runtime. Rasterizers, visual comparison tools, and external renderers belong in tests or development tooling.
- Polished invoice, report, and statement examples belong in samples and visual fixtures, not as special engine concepts.
- Adapter-specific mapping belongs in the source adapter packages. Shared PDF layout, reading, and manipulation behavior belongs here.
- Current-state inventories belong in [Docs/officeimo.pdf.current-state.md](../Docs/officeimo.pdf.current-state.md), not in this NuGet README.

## Repository validation

The repository keeps the public contract, target frameworks, package dependency
shape, performance budgets, compliance proof, and rendered output under
separate gates:

```powershell
dotnet test OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj -c Release -f net8.0
dotnet test OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj -c Release -f net10.0
dotnet run --project OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net8.0 -- --verify-budgets
Build/Export-PdfComplianceProof.ps1 -Configuration Release -Framework net8.0
Build/Export-PdfVisualReviewGallery.ps1 -Configuration Release -Framework net8.0
```

Pixel baselines are strict when the installed Poppler major/minor version
matches the recorded renderer. A different renderer version still runs semantic
and page-count checks in ordinary local runs. Required-rasterizer and CI visual
gates fail on a version mismatch; release investigations can deliberately opt
into a cross-version comparison.

## Current state

The PDF engine is useful and broad, but it is still evolving. It has strong first-party coverage for common generated business documents, conservative read/manipulation workflows, password security, optional first-party certificate signing/validation, standards-compliant Fast Web View output, and bounded-payload stream saves. Advanced typography, difficult producer-specific preservation, broader transparency/pattern edge cases, and fully forward-only layout remain deeper roadmap areas.

For the full capability inventory and roadmap, read [Docs/officeimo.pdf.current-state.md](../Docs/officeimo.pdf.current-state.md).

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Drawing`. PDF parsing, writing, logical recovery, manipulation, forms, diagnostics, and preservation analysis are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
