# OfficeIMO.Pdf - First-party PDF engine

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf)](https://www.nuget.org/packages/OfficeIMO.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Pdf)

`OfficeIMO.Pdf` is the first-party PDF package for OfficeIMO. It creates, reads, inspects, edits, merges, splits, stamps, exports, signs, and validates PDFs. PDF mechanics, password security, signature structure, and rendering remain first-party. Certificate-based CMS, RFC 3161, and X.509 operations use an explicitly supplied provider from the optional `OfficeIMO.Security` package.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Pdf
```

Install the security provider only when the application creates or cryptographically validates certificate-based PDF
signatures:

```powershell
dotnet add package OfficeIMO.Security
```

## Quick start

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("OfficeIMO.Pdf")
        .Paragraph(p => p
            .Text("A first-party PDF builder with ")
            .Bold("rich text")
            .Text(", links, tables, images, and document operations."))
        .Table(new[] {
            new[] { "Area", "Status" },
            new[] { "Optional signature provider", "OfficeIMO.Security" },
            new[] { "License", "MIT" }
        })), new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 11
        })
    .Meta(title: "Hello PDF", author: "OfficeIMO")
    .Save("hello.pdf");
```

## What it does

- Creates PDFs with page setup, headings, paragraphs, rich text, links, lists, reusable typed and page-aware components, tested report/invoice/label-sheet/ticket recipes, mixed inline images and boxes, dictionary-driven hyphenation, styled multipage containers, balanced block-flow columns, conditional/replayable flow, position capture, sections, generated TOCs, optional-content layers, tables, images, vector drawing, headers, footers, watermarks, metadata, portfolios, and form primitives. Raster inputs accepted by `OfficeIMO.Drawing` normalize once through the shared image owner before PDF embedding.
- Reads and inspects PDFs through text extraction, logical document objects, page metadata, links, images, attachments, portfolios, outlines, forms, bounded immutable raw-structure views, active-content diagnostics, and security/revision markers.
- Manipulates existing PDFs with page extraction, split, merge, delete, duplicate, move, rotate, metadata editing, stamps, watermarks, and complete-page overlay/underlay while preserving source PDF header versions on shared rewrite paths.
- Renders supported embedded TrueType and OpenType/CFF fonts with stable-glyph subsetting. `UseManagedTextShaping()` selects Drawing's dependency-light positioned-glyph provider for its proven core-Arabic/TrueType subset. The shared `IOfficeTextShapingProvider` contract remains the extension point for broader scripts and shaping engines.
- Projects authored annotation appearance streams into page images. When a supported free-text, text-markup, shape, line, ink, path, stamp, or caret annotation has no usable normal appearance, the renderer reuses the bounded annotation synthesizer and reports `render.annotation.appearance-synthesized` as an approximation.
- Shares managed CMYK, Lab, XYZ, calibrated-color conversion, bounded sampled, exponential, stitching, and Type 4 calculator color functions, vector tiling fills, standard blend modes, and alpha/luminosity soft masks with `OfficeIMO.Drawing`. Catalog destination output profiles with supported RGB matrix/TRC or ICC mBA transforms soft-proof vector, text, form, pattern, and image colors through the same rendering-intent pipeline. ICC LUT-composed and output-profile-composed shadings remain fail-closed unless their final interpolation can be certified. Pages with explicit transparency retain authored colors and report `render.colorspace.icc-output-intent-transparency-simplified` until output conversion can run after composition. Color-managed DCT/JPEG images use the ICC, `/Decode`, Indexed-palette, and transparency pipeline, while simple device-color JPEGs without required color management remain lossless pass-through payloads.
- Bounds completed page/effect content and serialized-object retention with separate memory limits, temporary-file spillover, direct large-stream spooling, and chunked final assembly during stream saves. `PdfSaveResult.Serialization` records limits, peak retained bytes, spill decisions, final buffering, and passthrough without claiming forward-only layout. Per-page metadata and the authored block model remain proportional to document size, and `ToBytes()` buffers the final artifact.
- Provides conversion reports, grouped warning summaries, and diagnostics so adapters can expose unsupported or simplified source content honestly.
- Provides reusable conversion proof snapshots for generated PDFs, artifact hashes, required page counts, page sizes, document metadata, outline titles, URI links, form fields, named destinations, page labels, attachments, output intents, optional-content/layer metadata, catalog/viewer metadata, XMP/tagged metadata, text markers, logical readback signals, expected and accepted warning contracts, and post-processing hand-off. Compliance proof records bind external validator name, version, profile, result, warnings, SHA-256, byte length, and validation time to the exact artifact.
- Provides reusable rewrite-preservation proof for page geometry, metadata, navigation, catalog/viewer/action state, optional content, tagged content, security signatures, document versions, and source-structure markers such as incremental updates, xref streams, and object streams.
- Provides a reusable rewrite-preservation matrix for classifying named manipulation scenarios as rewrite-safe, preservation-failed, blocked by safety checks, or operation-failed, including optional-content/layer drift, targeted form-fill preservation, form/tagged/active-content/signature blockers, and fluent `PdfDocument` helpers for normal document rewrite operations.
- Serves as the shared engine for Word, Excel, PowerPoint, OpenDocument, Markdown, HTML, RTF, OneNote, AsciiDoc, and LaTeX PDF adapters.

## Existing PDF workflows

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("input.pdf")
    .Pages.Extract("1-2,4")
    .MergeWith("appendix.pdf")
    .UpdateMetadata(title: "Merged report")
    .Stamp.Text("Reviewed")
    .Save("output.pdf");

PdfDocumentReadResult read = PdfDocument.Load("output.pdf").Read();
string text = string.Join('\n', read.Pages.SelectMany(page => page.TextBlocks).Select(block => block.Text));
```

`Load(...)` is the one entry point for byte arrays, files, and streams. It
enforces the same `PdfLoadOptions` limits before buffering, snapshots caller
input once, and reuses one parsed document across read, inspection, preflight,
diagnostic, optimization, signature, and compliance operations.

### Read and edit named document JavaScript

```csharp
using OfficeIMO.Pdf;

PdfDocument document = PdfDocument.Load("input.pdf");

foreach (PdfJavaScript script in document.JavaScript.List()) {
    Console.WriteLine(script.Name);
}

PdfJavaScriptEditResult edited = document.JavaScript.Edit(scripts => scripts
    .AddOrReplace("Initialize", "this.zoom = 100;")
    .Remove("Obsolete"));

File.WriteAllBytes("output.pdf", edited.ToBytes());
```

Script names use exact, case-sensitive matching. Editing preserves untouched
name-tree entries and action data, then reads the saved artifact back before
returning it. Per-script, script-count, and aggregate-byte limits come from
`PdfLoadOptions.Limits`. Document JavaScript is active content: the default
sanitizer removes it, and full-rewrite edits are blocked for encrypted or signed
inputs rather than weakening their security or revision contracts.

### Preview and select active-content removal

```csharp
PdfDocument incoming = PdfDocument.Load("incoming.pdf");
var policy = new PdfSanitizationOptions {
    ActionKindsToRemove = PdfSanitizationActionKind.JavaScript |
        PdfSanitizationActionKind.Launch |
        PdfSanitizationActionKind.SubmitForm
};

PdfSanitizationReport preview = incoming.InspectSanitization(policy);
Console.WriteLine($"Scripts: {preview.ActionCounts.JavaScript}");
Console.WriteLine($"Launch actions: {preview.ActionCounts.Launch}");

PdfSanitizationResult result = incoming.Sanitize(policy);
File.WriteAllBytes("sanitized.pdf", result.ToBytes());
```

`ActionKindsToRemove` is an exact opt-in selection, so unselected action kinds
remain. Selecting `Uri` removes every URI action and catalog URI base, including
ordinary web links. Leave the property null for the established default policy:
known active-content actions are removed, allowed `http`, `https`, `mailto`, and
`tel` links remain, and URI schemes outside `AllowedUriSchemes` are removed.

### Inspect and sanitize before sharing

```csharp
PdfDocument incoming = PdfDocument.Load("incoming.pdf");
var policy = new PdfSanitizationOptions {
    ContentKindsToRemove = PdfSanitizationContentKind.All,
    ActionKindsToRemove = PdfSanitizationActionKind.All
};

PdfSanitizationReport preview = incoming.InspectSanitization(policy);
Console.WriteLine($"User metadata: {preview.CategoryCounts.UserMetadata}");
Console.WriteLine($"Attachments: {preview.CategoryCounts.EmbeddedFiles}");
Console.WriteLine($"Actions: {preview.CategoryCounts.Actions}");
Console.WriteLine($"Comments and markup: {preview.CategoryCounts.CommentsAndMarkup}");
Console.WriteLine($"Bookmarks: {preview.CategoryCounts.Bookmarks}");
Console.WriteLine($"Layer definitions: {preview.CategoryCounts.OptionalContent}");

PdfSanitizationResult result = incoming.Sanitize(policy);
File.WriteAllBytes("shared.pdf", result.ToBytes());
```

`ContentKindsToRemove` is an exact selection of user-authored Info fields and
XMP, embedded files, actions, comments and markup, bookmarks, and optional-content
definitions. Unselected categories remain. Link and Widget annotations remain
when comments are selected; selecting URI actions can still remove a link's URI.
Producer, creation and modification dates, and trapping status are not classified
as user metadata and remain in the output.

Selecting optional content removes layer definitions and associations. Drawing
and text operators that belonged to a layer remain as ordinary page content, so
they no longer depend on viewer layer state. Use content-safety inspection and
verified redaction when the page content itself must be removed.

### Add interactive fields to an existing PDF

```csharp
PdfAcroFormEditResult edited = PdfDocument.Load("input.pdf").Forms.Edit(form => form
    .Create(new PdfFormFieldCreateOptions {
        Name = "customer.notes",
        Kind = PdfFormFieldCreationKind.Text,
        PageNumber = 1,
        X = 72,
        Y = 560,
        Width = 240,
        Height = 60,
        Style = new PdfFormFieldStyle { IsMultiline = true }
    })
    .Create(new PdfFormFieldCreateOptions {
        Name = "calculate",
        Kind = PdfFormFieldCreationKind.PushButton,
        PageNumber = 1,
        X = 72,
        Y = 520,
        Width = 100,
        Height = 24,
        Caption = "Calculate",
        JavaScript = "this.getField('total').value = 42;"
    }));

File.WriteAllBytes("form.pdf", edited.ToBytes());
```

The same transaction creates text fields, check boxes, combo or list choices,
radio-button groups, push buttons, and empty signature fields. Generated widget
appearances use `PdfFormFieldStyle`; widget JavaScript is returned as inert,
typed `PdfFormWidgetAction` data and is never executed by OfficeIMO.Pdf. Form
edits and fills preserve eligible actions owned by form widgets. Unrelated
catalog, page, outline, or annotation active content remains a rewrite blocker.
Flattening removes actions owned by the fields being flattened, while the
sanitizer removes forbidden actions without discarding the remaining form tree.

For a single health and capability view:

```csharp
PdfAnalysisReport analysis = PdfDocument
    .Load("incoming.pdf")
    .Analyze(PdfComplianceProfile.PdfA2B);

Console.WriteLine($"Pages: {analysis.Info.PageCount}");
Console.WriteLine($"Readable: {analysis.CanRead}");
Console.WriteLine($"Rewrite safe: {analysis.CanRewrite}");
Console.WriteLine($"Healthy: {analysis.IsHealthy}");

foreach (PdfDiagnosticFinding finding in analysis.Diagnostics.Findings) {
    Console.WriteLine($"{finding.Severity}: {finding.Code} — {finding.Message}");
}
```

## Unified document API

The unified API intentionally narrows the public surface around the fluent
`PdfDocument` facade:

- Use `PdfDocument.Load(...)` for normal existing-document workflows. Use
  `PdfReadDocument.Open(...)` only when a low-level parser model is the intended
  result rather than the canonical semantic document model.
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
- Author new documents through `PdfDocument.Create(pdf => ...)` or append authored
  content through `Compose(...)`. Headings, paragraphs, tables, images, and other
  flow primitives live on `PdfItemCompose`; they are no longer duplicated on the
  root document.
- Use the fluent `Pages`, `Forms`, `Attachments`, `Bookmarks`, `Annotations`,
  `Stamp`, `Security`, `Redactions`, `Optimization`, `Proof`, and metadata operations instead of the former public
  static engine classes. Those implementation engines are now internal so there
  is one supported route for each operation.
- `Save(...)`, `SaveAsync(...)`, and every typed adapter `SaveAsPdf(...)` now
  return `PdfSaveResult`. It carries output path/length, conversion warnings,
  and an immutable `Pipeline` with create/open, mutation, hash, page-count,
  execution-mode, timing, and final-output evidence. `TrySave(...)` keeps the
  same result shape while capturing exceptions instead of throwing.

The target-framework support remains `netstandard2.0`, `net8.0`, and
`net10.0`. See the [OfficeIMO migration guide](https://github.com/EvotecIT/OfficeIMO/blob/master/MIGRATION.md)
for the OfficeIMO 3.3 breaking changes and old-to-new API map.

## Examples

### Export PDF pages as images

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

PdfReadDocument pdf = PdfReadDocument.Open("input.pdf");

pdf.Pages[0]
    .ToImage()
    .AtDpi(144)
    .AsThumbnail(800)
    .AsPng()
    .Save("preview.png");

pdf.ToImages()
    .Pages("1-3,last")
    .FitWithin(1600, 1200)
    .WithMaximumRasterPixels(20_000_000)
    .AsWebp()
    .Save("page-images");

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Authored PDF")
        .Paragraph(paragraph => paragraph.Text("The authored model uses the same page renderer."))))
    .ToImages()
    .AsPng()
    .Save("authored-page-images");
```

PNG, JPEG, TIFF, SVG, and WebP use the same `OfficeImageExportResult` contract and Drawing-owned encoders. Pixel-fit limits apply consistently to vector and raster output, and allocation limits are resolved before a raster buffer is created. Unsupported or simplified PDF operators and resources remain visible as typed image diagnostics.

Any adapter that returns `PdfDocumentConversionResult` can use the same paged-image bridge without adding another renderer:

```csharp
IReadOnlyList<OfficeImageExportResult> pages = markdown
    .ToPdfDocumentResult()
    .ToImages()
    .AsPng()
    .Export();
```

Source conversion warnings are copied into every page result. Use `PdfReadPage.ToDrawing()` only when an intermediate `OfficeDrawing` is needed.

### Write a generated PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Page(page => page
        .Header(h => h.AlignCenter().Text("Service report"))
        .Footer(f => f.AlignRight().Text("Page {page} of {pages}"))
        .Content(layout => layout.Item(content => content
            .H1("Service report")
            .Paragraph(p => p
                .Text("Generated ")
                .Bold(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm 'UTC'"))
                .Text(" with first-party PDF primitives."))
            .Table(new[] {
                new[] { "System", "Status", "Owner" },
                new[] { "Identity", "Green", "Operations" },
                new[] { "Messaging", "Yellow", "Exchange" }
            })))), new PdfOptions {
            PageSize = PageSizes.A4,
            Margins = PageMargins.UniformCentimeters(1.6),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        })
    .Meta(
        title: "Service report",
        author: "OfficeIMO",
        subject: "Generated PDF")
    .Save("service-report.pdf");
```

Generated headers and footers can combine literal text, visually styled runs, and styled page tokens. The same builder is available for the default, first-page, and even-page variants:

```csharp
PdfDocument.Create(pdf => pdf.Page(page => page
    .Header(header => header
        .Text(text => text
            .Run(PdfTextRun.Bolded("Confidential ", PdfColor.FromRgb(180, 0, 0)))
            .Text("- page ")
            .CurrentPage(PdfTextRun.Italicized(string.Empty))
            .Text(" of ")
            .TotalPages(PdfTextRun.Italicized(string.Empty)))
        .FirstPageText(text => text.Run(PdfTextRun.Bolded("Confidential cover")))
        .EvenPagesText(text => text.Run(PdfTextRun.Underlined("Confidential even page"))))
    .Content(layout => layout.Item(content => content
        .Paragraph(p => p.Text("Generated report body."))))))
    .Save("styled-header.pdf");
```

Styled header/footer runs support fonts, size, color, highlighting, underline, strike, and baseline changes. Use the existing header/footer image and shape methods for visuals; interactive links and inline elements are intentionally kept out of text runs. Authored header/footer content that enters margins or overlaps another zone is preserved instead of rejected; attach a `PdfConversionReport` with `ReportDiagnosticsTo(...)` when the host needs structured overflow or clipping warnings.

### Rich report layout

```csharp
PdfDocument.Create(pdf => pdf.Content(content => content
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
        })))
    .Save("summary.pdf");
```

### Reusable business recipes

```csharp
var invoice = new PdfInvoiceComponent(
    invoiceNumber: "INV-42",
    issueDate: DateTime.Today,
    seller: new PdfInvoiceParty("Seller Ltd", new[] { "Tax ID 123" }),
    customer: new PdfInvoiceParty("Customer Ltd"),
    lines: new[] { new PdfInvoiceLine("Engineering", 2M, 50M, taxRate: 0.20M) },
    currencyCode: "EUR");

PdfDocument.Create(pdf => pdf.Content(content => content
        .Component(new PdfReportComponent("Delivery summary", "All checks passed."))
        .Component(invoice)))
    .Save("delivery-pack.pdf");
```

These recipes compose normal flow, table, and panel primitives. `IPdfContextComponent`
uses the existing deferred replay path when content must react to the live page number;
it does not introduce another layout engine.

### Hyphenation and inline visuals

```csharp
byte[] statusIcon = File.ReadAllBytes("status.png");
var hyphenation = new PdfHyphenationLexicon(new[] {
    "auto-ma-tion",
    "ty-pog-ra-phy",
    "re-port-ing"
});

PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(paragraph => paragraph
            .Text("Automation status ")
            .InlineImage(statusIcon, 12, 12, alternativeText: "Healthy")
            .Text(" remains available during long reporting runs."))),
    new PdfOptions().UseTextHyphenationDictionary(hyphenation))
    .Save("inline-status.pdf");
```

Inline elements participate in normal line wrapping. In tagged output, image and box alternative text is carried into the structure tree.

### Sections, generated navigation, and bounded stream output

```csharp
var options = new PdfOptions {
    PageContentMemoryLimitBytes = 4 * 1024 * 1024,
    ObjectBufferMemoryLimitBytes = 8 * 1024 * 1024
};

PdfSaveResult save = PdfDocument.Create(pdf => pdf.Content(content => content
        .TableOfContents()
        .Section("Summary", section => section
            .Container(container => container
                .Paragraph(p => p.Text("A styled, keep-together summary."))))
        .Section("Details", section => section
            .Columns(columns => {
                columns.Paragraph(p => p.Text("First column"));
                columns.ColumnBreak();
                columns.Paragraph(p => p.Text("Second column"));
            }, new PdfMultiColumnOptions { ColumnCount = 2, Gap = 18 }))), options)
    .Save("navigable-report.pdf");

Console.WriteLine($"Peak page payload: {save.Serialization?.PeakRetainedPageContentBytes}");
Console.WriteLine($"Object spill used: {save.Serialization?.ObjectBufferSpilled}");
```

### Load once and build one semantic read result

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("statement.pdf");
PdfDocumentReadResult result = pdf.Read(); // Structured is the default profile.

foreach (PdfLogicalPage page in result.Pages) {
    foreach (PdfLogicalParagraph paragraph in page.Paragraphs) {
        Console.WriteLine(paragraph.Text);
    }

    foreach (PdfLogicalTable table in page.Tables) {
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Console.WriteLine($"Page {page.PageNumber}: {data.Rows.Count} table rows");
    }
}

string markdown = result.ToMarkdown();
PdfMetadata metadata = result.Metadata;
PdfDocumentSecurityInfo security = result.Security;
IReadOnlyList<PdfOutlineItem> outlines = result.Outlines;
IReadOnlyList<PdfLogicalLinkAnnotation> links = result.Links;
IReadOnlyList<PdfFormField> formFields = result.FormFields;
IReadOnlyList<PdfAttachmentInfo> attachmentMetadata = result.Attachments;

IReadOnlyList<PdfExtractedImage> images = pdf.Images.Extract();
IReadOnlyList<PdfImagePlacement> placements = pdf.Images.Placements("1-2");
IReadOnlyList<PdfExtractedAttachment> attachments = pdf.Attachments.Extract();
```

`PdfDocument.Read(...)` is the only semantic reconstruction entry point. Both
profiles return `PdfDocumentReadResult`; they do not maintain separate logical
models. `Structured` adds document-wide tagged-PDF, outline, repeated-edge,
heading-tier, hierarchy, and continuation evidence. `Fast` keeps the same page
pipeline and result contract while omitting optional document-wide enrichment:

```csharp
PdfDocumentReadResult selected = pdf.Read(new PdfReadOptions {
    Profile = PdfReadProfile.Fast,
    PageSelection = PdfPageSelection.Parse("1-2,5")
});
```

Every built-in semantic stage observes the supplied cancellation token while it
works, not only between pages. Custom pipeline stages receive the same token and
work-budget context. `PdfUnderstandingPipelineOptions.MaxWorkUnitsPerPage` and
`MaxDocumentWorkUnits` bound geometry comparisons, recursive partitioning, and
document-wide evidence; over-complex input fails with
`PdfReadLimitKind.UnderstandingWork` instead of continuing unbounded work.

### Inspect fonts, image paint, and cross-page continuations

The font inventory reports every unique declared font dictionary and every page
or nested Form XObject resource path that references it. Embedded OpenType and
TrueType programs are parsed through the same bounded font engine used by PDF
generation. The parsed glyph count describes the embedded program; it is not a
count of glyphs painted by the page.

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("annual-report.pdf");
PdfDocumentReadResult result = pdf.Read();
PdfFontInventory fonts = pdf.Resources.Fonts();

foreach (PdfFontInfo font in fonts.Fonts) {
    Console.WriteLine($"{font.FamilyName}: embedded={font.IsEmbedded}, subset={font.IsSubset}");
    if (font.EmbeddedOpenTypeInfo is PdfOpenTypeFontInfo program) {
        Console.WriteLine($"Program glyphs: {program.GlyphCount}; Unicode mappings: {program.UnicodeScalarCount}");
    }
}

IReadOnlyList<PdfImagePlacement> placements = pdf.Images.Placements();
foreach (PdfImagePlacement placement in placements) {
    Console.WriteLine(
        $"Page {placement.PageNumber}: {placement.Width} x {placement.Height}, " +
        $"opacity={placement.Opacity}, blend={placement.EffectiveBlendMode}, clipped={placement.Clip is not null}");
}

IReadOnlyList<PdfLogicalParagraphContinuationGroup> paragraphs =
    result.GetParagraphContinuationGroups(new PdfLogicalParagraphContinuationOptions {
        MinimumConfidence = 0.80
    });
IReadOnlyList<PdfLogicalTableContinuationGroup> tables =
    result.GetTableContinuationGroups(new PdfLogicalTableContinuationOptions {
        MinimumConfidence = 0.80
    });

foreach (PdfLogicalParagraphContinuationGroup paragraph in paragraphs.Where(item => item.SpansPages)) {
    Console.WriteLine($"Paragraph {paragraph.FirstPageNumber}-{paragraph.LastPageNumber}: {paragraph.Confidence:P0}");
}
foreach (PdfLogicalTableContinuationGroup table in tables.Where(item => item.SpansPages)) {
    Console.WriteLine($"Table {table.FirstPageNumber}-{table.LastPageNumber}: {table.TotalRowCount} rows");
}
```

Cross-page recovery is conservative. It joins only adjacent boundary items with
compatible geometry and content evidence, returns confidence and evidence flags,
and leaves the original page-local logical model unchanged. Embedded program bytes
are returned only when `IncludeEmbeddedProgramBytes` is explicitly enabled and
remain subject to both `MaxEmbeddedProgramBytes` per font and
`MaxTotalDecodedFontBytes` across embedded programs and ToUnicode maps. ToUnicode
per-map and aggregate limit outcomes have separate diagnostics and inventory counts,
so resource limits are not reported as malformed maps. Nested Form XObject discovery
deduplicates shared resource contexts by their shallowest reachable depth and is
bounded by `MaxFormResourceTraversals`.
For image paint inspection, `AuthoredBlendMode` is null when the normal PDF default
was not declared and retains an explicit or inherited authored `Normal` value.

Text extraction excludes PDF artifact marked content by default, which is the
logical-text behavior expected for decorative headers, footers, and chart
labels. Opt into visual text when those marked artifacts are part of the
required payload:

```csharp
PdfDocument visualTextPdf = PdfDocument.Load("spreadsheet-export.pdf", new PdfLoadOptions {
    IncludeArtifactText = true
});

PdfDocumentReadResult visual = visualTextPdf.Read();
IReadOnlyList<string> visualTextByPage = visual.Pages
    .Select(page => string.Join('\n', page.TextBlocks.Select(block => block.Text)))
    .ToArray();
```

Image-only and mixed pages can be enriched through a caller-owned OCR provider
without adding an OCR runtime to `OfficeIMO.Pdf`. Accepted words are normalized
to cropped, rotated visual page coordinates, de-duplicated against native text,
and projected into the same logical model used by reverse converters:

```csharp
static async Task<PdfDocumentReadResult> ReadWithOcrAsync(
    string path,
    IPdfOcrProvider provider) {
    PdfOcrMergeResult ocr = await PdfDocument
        .Load(path)
        .Ocr.ReadAsync(provider, new PdfOcrMergeOptions {
            MinimumConfidence = 0.75,
            DetectAlignedTables = true
        });

    Console.WriteLine(ocr.AcceptedWordCount);
    return ocr.EnrichedDocument;
}
```

`NativeDocument` retains the parser-only view for comparison. Every enriched
text block and inferred table exposes native-or-OCR provenance; OCR text also
retains provider confidence and direct visual bounds. Table inference requires
repeated aligned rows plus typed-value evidence (or a wider repeated grid), so
ordinary two-column prose remains separate reading-order content. The work is
bounded by the merge options. Word, Excel,
PowerPoint, HTML, RTF, ODT, ODS, and ODP packages consume
`EnrichedDocument` directly through their existing logical-PDF overloads.

### Split and extract pages

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Load("packet.pdf");

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

PdfDocument.Load("packet.pdf")
    .MergeWith("appendix.pdf")
    .Pages.Delete("2,5-6")
    .Pages.Duplicate("1")
    .Pages.Move(insertBeforePageNumber: 3, pageRanges: "7-8")
    .Pages.Rotate(90, "4")
    .UpdateMetadata(title: "Cleaned packet")
    .Save("packet-clean.pdf");
```

Encrypted merge inputs keep independent authentication settings. Owner
authorization is honored automatically. A user password follows the PDF
permission bits unless the caller explicitly opts into ignoring those
restrictions:

```csharp
PdfDocument first = PdfDocument.Load("first.pdf", new PdfLoadOptions {
    Password = "first-owner-password"
});
PdfDocument second = PdfDocument.Load("second.pdf", new PdfLoadOptions {
    Password = "second-user-password",
    PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
});

PdfMergeResult merged = PdfDocument.MergeWithReport(
    new PdfMergeOptions(),
    first,
    second);

File.WriteAllBytes("merged.pdf", merged.ToBytes());
Console.WriteLine(merged.Report.OutputHasEncryption); // False
Console.WriteLine(merged.Report.Sources[1].PermissionRestrictionsIgnored); // True
```

`IgnoreRestrictions` is an authenticated permission override, not password
recovery. The document must still decrypt with the supplied password; an
unknown or incorrect password remains an error. Full rewrites of signed PDFs
remain blocked because they would invalidate existing signatures.

### Production document workflows

Apply one continuous Bates sequence across a batch. Each output includes its
page assignments and rewrite-preservation evidence:

```csharp
var inputs = new[] {
    new PdfBatesDocument(File.ReadAllBytes("volume-a.pdf"), "volume-a.pdf"),
    new PdfBatesDocument(File.ReadAllBytes("volume-b.pdf"), "volume-b.pdf")
};

PdfBatesBatchResult numbered = PdfBatesNumberer.Apply(inputs, new PdfBatesNumberingOptions {
    StartNumber = 1200,
    Prefix = "CASE-",
    MinimumDigits = 6,
    Position = PdfBatesPosition.BottomRight
});

File.WriteAllBytes("volume-a-numbered.pdf", numbered.Documents[0].ToBytes());
Console.WriteLine(numbered.NextNumber);
```

Every label is measured against its target page rectangle. A batch fails
instead of wrapping or clipping an identifier that cannot be rendered in full.

Interleave selected pages with explicit output provenance, including reverse
order for duplex scan backs:

```csharp
var fronts = new PdfInterleaveSource(File.ReadAllBytes("fronts.pdf"), "fronts");
var backs = new PdfInterleaveSource(File.ReadAllBytes("backs.pdf"), "backs") {
    Reverse = true
};

PdfInterleaveResult interleaved = PdfPageInterleaver.Interleave(
    new[] { fronts, backs },
    new PdfInterleaveOptions { RemainderMode = PdfInterleaveRemainderMode.Reject });

File.WriteAllBytes("duplex-scan.pdf", interleaved.ToBytes());
```

Page-addressed form fields and named links follow the selected pages and their
final output ownership; structures that belong only to excluded pages are not
carried into the composed document.

Production splitting can combine page-count, text-boundary, and target-size
rules. Every part reports its source pages, termination reason, byte size, and
whether one indivisible page exceeded the target. Probe count and cumulative
generated bytes are bounded and reported for the complete operation:

```csharp
PdfProductionSplitResult split = PdfProductionSplitter.Split(
    File.ReadAllBytes("records.pdf"),
    new PdfProductionSplitOptions {
        MaximumPagesPerPart = 100,
        BoundaryText = "START OF RECORD",
        TargetPartSizeBytes = 20L * 1024 * 1024
    });

foreach (PdfProductionSplitPart part in split.Parts) {
    File.WriteAllBytes($"records-{part.PartNumber:000}.pdf", part.ToBytes());
}
```

These operations reuse the existing selector, importer, stamping, mutation
preflight, and preservation engines. They do not bypass encryption,
permissions, signatures, or catalog-structure policy.

### Verified repair artifacts

Lenient parsing can persist only explicitly recovered structural defects into
a normalized artifact. The result must reopen in strict mode and pass the
rewrite-preservation comparison:

```csharp
PdfRepairArtifactResult repaired = PdfRepairArtifact.Create(
    File.ReadAllBytes("recovered-input.pdf"));

if (repaired.IsVerified) {
    File.WriteAllBytes("repaired.pdf", repaired.ToBytes());
}
```

Detected-only ambiguity is rejected by default. Encrypted PDFs, signatures,
certification permissions, and usage rights are not silently removed by the
repair workflow.

### Annotation review threads

Read reply relationships as threads, add a reply, and record a standard review
state through the same annotation mutation policy used by lower-level edits:

```csharp
byte[] source = File.ReadAllBytes("review.pdf");
PdfAnnotationReviewCatalog review = PdfAnnotationReviewCatalog.Read(source);
int rootObjectNumber = review.Threads[0].Root.Annotation.ObjectNumber!.Value;

PdfAnnotationEditResult reply = PdfAnnotationReviewEditor.AddReply(
    source,
    rootObjectNumber,
    "Confirmed against the source record.",
    new PdfAnnotationReplyOptions { Author = "Reviewer", Subject = "Evidence" });

PdfAnnotationReviewCatalog updated = PdfAnnotationReviewCatalog.Read(reply.Bytes);
int replyObjectNumber = updated.Threads[0].Root.Replies[0].Annotation.ObjectNumber!.Value;
PdfAnnotationEditResult accepted = PdfAnnotationReviewEditor.SetState(
    reply.Bytes,
    replyObjectNumber,
    PdfAnnotationReviewState.Accepted);
```

Replies expose raw `/IRT`, `/RT`, `/State`, `/StateModel`, `/Subj`, and `/IT`
metadata alongside the typed standard state. Certified documents use an
append-only revision only when the signature permission model allows the
annotation change. Thread construction has explicit relationship and nesting
limits so hostile reply chains fail closed.

### Password protection on browser or restricted hosts

Desktop and server applications use platform AES automatically. A host without synchronous platform AES can pass the
managed provider included in `OfficeIMO.Core` explicitly; the same provider handles writing and opening the protected PDF.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Security;

IOfficeAesCryptographyProvider aes = OfficeManagedAesCryptographyProvider.Default;
var encryption = new PdfStandardEncryptionOptions("reader-password") {
    OwnerPassword = "owner-password",
    Algorithm = PdfStandardEncryptionAlgorithm.Aes256,
    AesCryptographyProvider = aes
};

PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(p => p.Text("Protected in the current host."))),
        new PdfOptions().SetEncryption(encryption))
    .Save("protected.pdf");

PdfDocument opened = PdfDocument.Load("protected.pdf", new PdfLoadOptions {
    Password = "owner-password",
    AesCryptographyProvider = aes
});
```

### Certificate-based PDF signatures

PDF signature discovery, byte-range inspection, mutation blocking, and caller-defined external signing do not require
`OfficeIMO.Security`. For the built-in CMS adapter, install the optional package and pass its provider explicitly:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
using var signer = new PdfCmsExternalSigner(security, signingCertificate);

PdfExternalSignatureCompletion signed = PdfDocument
    .Load("contract.pdf")
    .Security.SignExternal(
        signer,
        new PdfExternalSignatureOptions {
            FieldName = "Approval",
            VisibleAppearance = new PdfVisibleSignatureAppearanceOptions {
                ImageBytes = File.ReadAllBytes("approval-mark.png"),
                ImageFit = OfficeImageFit.Contain,
                ShowText = false
            }
        });

var cryptography = new PdfCmsSignatureCryptographyProvider(
    security,
    new CmsVerificationOptions());
PdfSignatureValidationReport report = signed.ToDocument().Security.ValidateSignatures(cryptography);
```

The PDF package owns byte ranges, incremental updates, signature dictionaries, and preservation policy. The optional
provider owns CMS, timestamps, and certificate trust. A custom `IPdfExternalSigner` or
`IPdfSignatureCryptographyProvider` remains valid without `OfficeIMO.Security`.
The optional appearance image is visual content only; certificate validation remains the source of signer identity.

### Review, apply, and verify redactions

Build a source-bound plan, review its areas and matches, then apply that exact plan and retain the evidence report with the output:

```csharp
PdfDocument source = PdfDocument.Load("contract.pdf");
PdfRedactionPlan plan = source.Redactions.Search(
    new PdfRedactionSearchOptions().AddLiteral("Account: 123-45-6789"));

// Present plan.Areas and plan.Matches for approval before applying it.
var verification = new PdfRedactionVerificationOptions {
    RequireCompleteStreamInspection = true,
    CheckManagedRendering = true
}.RequireRemovedText("Account: 123-45-6789");

PdfRedactionApplyResult redacted = source.Redactions.ApplyWithEvidence(
    plan,
    verificationOptions: verification);

redacted.ThrowIfUnverified();
File.WriteAllBytes("contract-redacted.pdf", redacted.Pdf);
Console.WriteLine(redacted.Evidence.Summary);
```

`Evidence.Items` records a verified-absent, residual, or inconclusive outcome for every reviewed match. The report also exposes source/output hashes, residual matches, verification details, and affected page numbers. A UI can pass those page numbers to the existing page renderer for before/after previews without making rendering part of the redaction contract.

Text redaction rewrites native text-show operations at glyph granularity. Encoded glyphs outside the reviewed areas retain their original font resource and position; removed glyph advances become `TJ` displacements so adjacent text does not reflow. When a glyph mapping cannot be proven safe, the complete PDF text object is removed instead.

When `verificationOptions` is omitted, `ApplyWithEvidence` requires complete stream inspection and managed-rendering checks by default. Supply explicit options, as above, when the workflow also needs removed/retained markers or an external validator.

### Stamp and watermark an existing PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("contract.pdf")
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
PdfDocument.Load("contract.pdf")
    .Stamp.OverlayPage("letterhead.pdf", new PdfPageOverlayOptions {
        SourcePageNumber = 1,
        TargetPages = PdfPageSelector.Parse("all,!last"),
        Fit = PdfPageOverlayFit.Contain,
        Opacity = 0.9
    })
    .Save("contract-with-letterhead.pdf");
```

For richer existing-page automation, stamp a general visual canvas instead of
using separate table-, text-, and image-only operations:

```csharp
PdfDocument.Load("contract.pdf")
    .Stamp.Content((canvas, page) => {
        canvas.Text($"Page {page.PageNumber} of {page.PageCount}", 36, 24, 220, 24)
            .Table(new[] {
                new[] { PdfTableCell.TextCell("Status"), PdfTableCell.TextCell("Reviewed") },
                new[] { PdfTableCell.TextCell("Owner"), PdfTableCell.RichTextCell(new[] { PdfTextRun.Bolded("Legal") }) }
            }, 36, 620, page.Width - 72, 90);
    }, new PdfCanvasStampOptions {
        TargetPages = PdfPageSelector.Parse("1,last"),
        Opacity = 0.95
    })
    .Save("contract-with-review-panel.pdf");
```

Canvas stamping is intentionally visual-only. Text, rich tables, images,
shapes, drawings, clipping, and effects are supported. Interactive links and
annotations, named destinations, forms, and document outlines use their
dedicated editors so their behavior is not silently flattened or discarded.

### Search and edit existing page text

Text editing coordinates use PDF points from the page bottom-left. Inspect a
region when the UI needs the existing text and its detected style, then replace
or move it through the same text-removal and stamping owners used by redaction
and existing-page stamps:

```csharp
PdfDocument document = PdfDocument.Load("contract.pdf");
var region = new PdfPageRegion(pageNumber: 1, x: 72, y: 640, width: 260, height: 28);

PdfRegionText current = document.Text.Inspect(region);
Console.WriteLine($"{current.Text} ({current.SourceFont}, {current.FontSize} pt)");

PdfTextEditResult edited = document.Text.Replace(region, "Approved", new PdfTextEditOptions {
    Color = PdfColor.FromRgb(25, 110, 55)
});

edited.Document.Text.Add(
        new PdfPageRegion(1, 72, 600, 240, 24),
        "Reviewed by Legal",
        new PdfTextEditOptions { Font = PdfStandardFont.HelveticaBold, FontSize = 11 })
    .Document
    .Save("contract-edited.pdf");
```

`Text.Find(...)` supports case and whole-word filters over visible, unclipped
text, while `Text.ReplaceAll(...)` preserves unmatched source-span text and
keeps wide same-baseline runs such as columns independent. Edits fail closed
when an atomic PDF text object would require invisible or clipped text to be
recreated without its original rendering state.
Unmatched glyphs remain encoded in their original font; newly inserted replacement text uses the closest standard PDF font unless the caller selects one.
`PdfTextEditResult.Warnings` reports source-font substitutions that can change
metrics or letterforms.

Invisible OCR text stored with PDF text rendering mode 3 is opt-in for both
discovery and mutation. Use `IncludeTextRenderingMode3` to find it, then
`AllowTextRenderingMode3` to authorize an edit that preserves the invisible
rendering mode:

```csharp
PdfDocument scanned = PdfDocument.Load("scanned-contract.pdf");
var ocrSearch = new PdfTextSearchOptions {
    MatchCase = true,
    IncludeTextRenderingMode3 = true
};

PdfTextMatch ocrMatch = scanned.Text.Find("Account number", ocrSearch).Single();
PdfTextEditResult corrected = scanned.Text.Replace(
    ocrMatch,
    "Customer number",
    new PdfTextEditOptions { AllowTextRenderingMode3 = true });

corrected.Document.Save("scanned-contract-corrected.pdf");
```

The OCR opt-ins do not authorize clipping text modes or Type3 font glyph
programs, because those glyph programs can paint visible graphics independently
of the text rendering mode. One edit also cannot combine visible text and
rendering-mode-3 OCR text.

### Find and edit existing page images

Image placement coordinates also use PDF points from the page bottom-left.
Discover placements through the editor, then remove, replace, or move one exact
invocation without deleting overlapping text, paths, annotations, or unrelated
images:

```csharp
PdfDocument document = PdfDocument.Load("contract.pdf");
PdfImagePlacement logo = document.Images.Find(
    new PdfPageRegion(pageNumber: 1, x: 36, y: 720, width: 180, height: 60)).Single();

PdfImageEditResult updated = document.Images.Replace(
    logo,
    File.ReadAllBytes("new-logo.png"),
    new PdfImageEditOptions { Layer = PdfImageEditLayer.AboveExistingContent });

PdfImagePlacement replacement = updated.Document.Images.Find(
    new PdfPageRegion(1, 36, 720, 180, 60)).Single();

updated.Document.Images.Move(replacement, deltaX: 12, deltaY: -8)
    .Document
    .Save("contract-with-new-logo.pdf");
```

`Images.Add(...)` fits a new image to a page region. Replacement and movement
preserve position, size, and rotation when the source transform is portable;
callers explicitly choose whether rewritten content is above or behind existing
page content. The editor fails closed for ambiguous placements and for source
clipping, opacity, skew/reflection, unresolved transparency, image-mask, raw
payload, or inline-image semantics that cannot be reproduced safely. Exact
XObject removal remains available for rotated and skewed placements.

### Fill and flatten a PDF form

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("application-form.pdf")
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

PdfComplianceArtifact artifact = PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(paragraph => paragraph.Text("This artifact is ready for external validation."))), options)
    .Meta(title: "Archive copy")
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

PdfDeclaredComplianceClaimsReport declaredClaims = PdfDocument
    .Load(pdf)
    .AssessDeclaredComplianceClaims(new[] { validation });

if (!proof.CanClaimConformance || !declaredClaims.CanClaimAllDeclaredConformance) {
    throw new InvalidOperationException(proof.ExternalProofSummary);
}
```

Formal generation gates are available for PDF/A-2a/b/u, PDF/A-3a/b/u, PDF/A-4/4e/4f, PDF/UA-1, PDF/UA-2, PDF/X-1a:2003, PDF/X-4, Factur-X, and ZUGFeRD. `RequireCompliance(...)` rejects incomplete generation settings. PDF/X additionally inspects the complete serialized artifact before any bytes are returned or committed to a destination. A conformance claim still requires a passing external result for the same profile, SHA-256, and byte length; validators are build-time tools and are not runtime dependencies of `OfficeIMO.Pdf`.

### Generate a fail-closed PDF/X artifact

```csharp
using OfficeIMO.Pdf;

byte[] cmykProfile = File.ReadAllBytes("PSOcoated_v3.icc");
byte[] fontBytes = File.ReadAllBytes("SourceSerif4-Regular.otf");
var options = new PdfOptions()
    .ConfigurePdfX(
        PdfComplianceProfile.PdfX4,
        cmykProfile,
        "FOGRA51",
        PdfTrappingStatus.False)
    .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "Source Serif 4");

PdfComplianceArtifact artifact = PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(paragraph => paragraph.Text("Generated colors are converted through the CMYK print condition."))),
        options)
    .Meta(title: "Print-ready report")
    .CreateComplianceArtifact(PdfComplianceProfile.PdfX4);

File.WriteAllBytes("print-ready.pdf", artifact.ToBytes());
```

`ConfigurePdfX` requires a caller-selected CMYK output-device profile because the correct profile depends on the printing condition and profile redistribution terms. For example, the [ICC registry entry for PSOcoated_v3](https://registry.color.org/profile-registry/PSOcoated_v3) identifies it as FOGRA51 and permits use, embedding, and exchange while restricting redistribution. OfficeIMO requires the ICC `prtr` device class, validates the header, declared size, CMYK component count, and a supported output transform, and writes a boolean trapping status (`False` by default). It also creates synchronized Info/XMP production dates, PDF/X identification, document and instance UUIDs, version identity, rendition class, and trapping metadata. It converts generated vector, text, and supported raster colors, applies the selected black-preservation policy, writes production page boxes, and inspects the exact PDF for remaining DeviceRGB content, embedded fonts, prohibited references, and profile-specific transparency. PDF/X-1a:2003 also rejects device-independent color spaces and flattens raster alpha before rejecting any remaining transparency. PDF/X-4 retains its standard color-management and transparency allowances, while OfficeIMO deliberately emits a conservative CMYK generated-content subset.

For reproducible builds, replace the generated timestamps and UUIDs with an explicit `PdfXProductionMetadata` value through `SetPdfXProductionMetadata(...)`. Reusing the same value produces stable production metadata; create a new document and instance UUID when the output represents a different resource.

Internal readiness is not a certification. Pass the exact artifact to a qualified PDF/X preflight tool and bind its result with `PdfExternalValidationResult.PassedForArtifact`; `PdfComplianceProofReport.CanClaimConformance` remains false when that exact external evidence is absent or mismatched.

### Choose converter-friendly text fallbacks

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("proposal.docx");

var options = new WordPdfSaveOptions {
    TextFallbacks = PdfTextFallbackFeatures.Default,
    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost()
}.UseProfile(PdfExportProfile.PrintReady);

var result = document.ToPdfDocumentResult(options);
result.Report.RequireNoErrorWarnings();
result.Save("proposal.pdf");
```

The Word, Excel, PowerPoint, Markdown, HTML, RTF, OneNote, AsciiDoc, and LaTeX PDF adapters use one `PdfResourcePolicy`; semantic-projection adapters expose it through their nested Markdown PDF options. The balanced default enables installed fonts and bounded data URI/package resources for document fidelity while denying arbitrary local files and remote resolver calls. Use `PdfResourcePolicy.CreatePortableDeterministic()` for reproducible or untrusted conversion, and `CreateTrustedHost()` only when both source and host are trusted. Profiles never grant resource access.

The text-capable adapters also expose `TextFallbacks`. `PdfTextFallbackFeatures.Default` enables document, monospace, symbol, and emoji groups. Add `PdfTextFallbackFeatures.MultilingualFonts` for CJK, Arabic, and other non-Latin family candidates; OneNote adds that candidate group unless fallbacks are `None`. Candidate selection does not read installed fonts unless the resource policy allows it.

### Inspect and remove content provenance

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;

OfficeProvenanceReport report = PdfProvenance.InspectFile("input.pdf");
OfficeProvenanceRemovalResult result = PdfProvenance.RemoveFile("input.pdf", "clean.pdf");
```

The PDF owner recognizes the standards-defined embedded file with media type `application/c2pa` and `/AFRelationship /C2PA_Manifest`. Removal uses a bounded, targeted object-graph rewrite that leaves unrelated attachment associations in place. Malformed indirect-object candidates remain untouched unless the caller disables `RequireStructurallyValidCarrier`; direct file-spec dictionaries are inspected but cannot be removed safely. A signed PDF is never silently rewritten or stripped; handle its signature through an explicit PDF signature workflow first. Optional cryptographic C2PA verification is provided by `OfficeIMO.Security`.

## Concealed-content inspection and cleanup

`PdfDocument.InspectContentSafety(...)` reports non-painting text render modes, effective transparency, clipping, tiny/zero/off-canvas geometry, paint-order-resolved low contrast, hidden annotation and form-widget values, text inside layers hidden by the default optional-content configuration, and Unicode evidence from decoded spans. `PdfDocument.RemoveSelectedContent(...)` physically removes exact selected spans and can rewrite reviewed Unicode ranges in ordinary painted or painted low-contrast spans while verifying neighboring text restoration. Hidden layer, annotation, and widget findings remain report-only because their container semantics cannot be safely converted into an exact text edit. Encrypted and signed PDFs are rejected.

### Generate a formal e-invoice carrier

```csharp
using OfficeIMO.Pdf;

byte[] invoiceXml = File.ReadAllBytes("factur-x.xml");
DateTimeOffset invoiceModifiedAt = File.GetLastWriteTimeUtc("factur-x.xml");
byte[] fontBytes = File.ReadAllBytes("SourceSerif4-Regular.otf");

PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(paragraph => paragraph.Text("Invoice preview"))),
    new PdfOptions()
        .UseFacturX(
            invoiceXml,
            relationship: PdfAssociatedFileRelationship.Alternative,
            textFallbacks: PdfTextFallbackFeatures.None)
        .SetEmbeddedFileModificationDate("factur-x.xml", invoiceModifiedAt)
        .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "Source Serif 4")
        .RequireCompliance(PdfComplianceProfile.FacturX))
    .Save("invoice.pdf");
```

The XML must be a valid EN 16931 CrossIndustryInvoice payload. The formal carrier gate checks the PDF/A-3 attachment, metadata, font, Unicode, and invoice rules before writing; exact-artifact PDF/A and invoice-validator results are still required before claiming conformance.

### Page setup, watermarks, and metadata

```csharp
PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Draft report")
        .Paragraph(paragraph => paragraph.Text("This document uses page-level options instead of post-processing."))),
    new PdfOptions {
            PageSize = PageSize.FromCentimeters(21, 29.7).Portrait(),
            Margins = PageMargins.UniformCentimeters(1.5),
            TextWatermark = new PdfTextWatermark("DRAFT") {
                Opacity = 0.12,
                RotationAngle = -35
            }
        })
    .Meta(title: "Draft report", author: "OfficeIMO")
    .Save("draft.pdf");
```

### Inspect and preflight before rewriting

```csharp
using OfficeIMO.Pdf;

byte[] bytes = File.ReadAllBytes("incoming.pdf");
PdfDocument pdf = PdfDocument.Load(bytes);
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
PdfDocument pdf = PdfDocument.Load("incoming.pdf");

var inspection = pdf.Inspect();
Console.WriteLine($"Pages: {inspection.PageCount}");
Console.WriteLine($"Links: {inspection.LinkAnnotationCount}");
Console.WriteLine($"Forms: {inspection.FormFields.Count}");
Console.WriteLine($"Active content: {inspection.HasActiveContent}");

foreach (var page in inspection.Pages) {
    Console.WriteLine($"{page.PageNumber}: {page.Width} x {page.Height}");
}

PdfMutationPortfolioReport mutations = pdf.AssessMutations();
PdfRenderCompatibilityReport rendering = pdf.AssessRenderCompatibility();
Console.WriteLine($"Executable mutation families: {mutations.ExecutablePlans.Count}");
Console.WriteLine($"Render capability findings: {rendering.DiagnosticCount}");
```

### Convert PDFs through adapter packages

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var word = WordDocument.Load("proposal.docx");
word.SaveAsPdf("proposal.pdf");

PdfDocument statement = PdfDocument.Load("bank-statement.pdf");
PdfExcelTableImportReport tableReport = statement.SaveTablesAsExcel(
    "bank-statement-tables.xlsx");

Console.WriteLine($"Non-table page content detected: {tableReport.HasOmittedPageContent}");

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
| [OfficeIMO.Excel.Pdf](../OfficeIMO.Excel.Pdf/README.md) | Maps Excel workbooks to PDF and recovers detected PDF tables as editable worksheets. |
| [OfficeIMO.Markdown.Pdf](../OfficeIMO.Markdown.Pdf/README.md) | Maps Markdown documents into PDF primitives. |
| [OfficeIMO.PowerPoint.Pdf](../OfficeIMO.PowerPoint.Pdf/README.md) | Maps PowerPoint slides to PDF and imports PDF pages as visual slides or editable detected tables. |
| [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) | Bridges HTML to PDF and PDF to HTML. |
| [OfficeIMO.Rtf.Pdf](../OfficeIMO.Rtf.Pdf/README.md) | Maps semantic RTF into PDF and logical PDF content back to RTF. |
| [OfficeIMO.OneNote.Pdf](../OfficeIMO.OneNote.Pdf/README.md) | Explicitly projects offline OneNote hierarchy into a semantic PDF document with loss diagnostics. |
| [OfficeIMO.AsciiDoc.Pdf](../OfficeIMO.AsciiDoc.Pdf/README.md) | Projects native AsciiDoc through the loss-aware Markdown bridge and combines parser, projection, and PDF diagnostics. |
| [OfficeIMO.Latex.Pdf](../OfficeIMO.Latex.Pdf/README.md) | Projects the bounded LaTeX profile through the loss-aware Markdown bridge without executing TeX. |
| [OfficeIMO.OpenDocument.Odt.Pdf](../OfficeIMO.OpenDocument.Odt.Pdf/README.md) | Provides direct ODT and PDF façades with combined OpenDocument and PDF diagnostics. |
| [OfficeIMO.OpenDocument.Ods.Pdf](../OfficeIMO.OpenDocument.Ods.Pdf/README.md) | Provides direct ODS and PDF façades without Word or PowerPoint dependencies. |
| [OfficeIMO.OpenDocument.Odp.Pdf](../OfficeIMO.OpenDocument.Odp.Pdf/README.md) | Provides direct ODP and PDF façades without Word or Excel dependencies. |

The generated [PDF conversion support matrix](../Docs/officeimo.pdf-conversion-support-matrix.md) records current direct and composed routes from the canonical [`Docs/pdf-conversion-scenarios.json`](../Docs/pdf-conversion-scenarios.json) manifest. `OfficeIMO.Pdf` projects the dependency-free `OfficeDocumentModel` through `PdfProjectionOptions`; `OfficeIMO.Reader.Pdf` keeps a thin compatibility bridge for existing `OfficeDocumentReadResult` workflows. Email, EPUB, and Visio are advertised as direct conversion only when their route-specific artifact gates are proven. Open PDF work is tracked in the repository [roadmap](../Docs/ROADMAP.md).

## Related packages and docs

- `OfficeIMO.Pdf` provides first-party PDF parsing, layout, writing, rendering, password security, and signature structure. Optional CMS, DER, and X.509 services come from an explicitly supplied `OfficeIMO.Security` provider.
- Source-format adapters map their document models onto the neutral `OfficeDocumentModel`; PDF projection remains owned by this package.
- See the [PDF current-state guide](../Docs/officeimo.pdf.current-state.md) for the detailed capability inventory and known limits.

## Repository validation

The repository keeps the public contract, target frameworks, package dependency
shape, performance budgets, compliance proof, and rendered output under
separate gates:

```powershell
dotnet test OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj -c Release -f net8.0
dotnet test OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj -c Release -f net10.0
dotnet run --project OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net8.0 -- --verify-budgets
dotnet run --project OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net10.0 -- --verify-budgets
dotnet run --project OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net10.0 -- --verify-timing-budgets
Build/Test-PdfQualityCorpus.ps1 -Configuration Release -Framework net8.0
Build/Test-RealWorldCorpusContract.ps1
Build/Export-PdfComplianceProof.ps1 -Configuration Release -Framework net8.0
Build/Export-PdfVisualReviewGallery.ps1 -Configuration Release -Framework net8.0
```

The strict PDF quality scorecard runs every public inspection and semantic stage,
the complete 21-operation mutation portfolio, managed rendering, and declared
compliance claim gating against hash-pinned Open Preservation Foundation and
veraPDF fixtures. The real-world corpus lane applies the same deep PDF stages to
a deterministic sample of a checksum-pinned public GovDocs archive in isolated
processes. The performance gate
uses a deterministic 60-page mixed corpus and checks cold and cached analysis,
SVG rendering, PNG rendering, output integrity, absolute allocation and heap
budgets, generous elapsed-time ceilings, and cached allocation savings. Relative
cached speedup is opt-in through `--verify-timing-budgets` for controlled
benchmark hosts; ordinary CI records it without treating shared-runner timing as
a release comparison.

Pixel baselines are strict when the installed Poppler major/minor version
matches the recorded renderer. A different renderer version still runs semantic
and page-count checks in ordinary local runs. Required-rasterizer and CI visual
gates fail on a version mismatch; release investigations can deliberately opt
into a cross-version comparison.

## Current state

The PDF engine is useful and broad, but it is still evolving. It has strong first-party coverage for common generated business documents, reusable Unicode line breaking and Latin ligatures, bounded built-in core-Arabic shaping plus an optional HarfBuzz adapter for full GSUB/GPOS shaping, and bounded Type 3 rendering within the documented capability contract. It also supports authored and bounded-synthesized annotation appearances in page images, conservative read/manipulation workflows, password security, optional provider-backed certificate signing/validation, standards-compliant Fast Web View output, and bounded-payload stream saves with runtime serialization evidence. See the [image export capability matrix](../Docs/officeimo.image-export-capability-matrix.md) for the exact Type 3 rendering coverage and current limitations.

For the current capability inventory, ownership boundaries, premium conversion
contract, and remaining general engine work, read
[Docs/officeimo.pdf.current-state.md](../Docs/officeimo.pdf.current-state.md).

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** no third-party PDF parser, writer, renderer, or cryptographic package in the base PDF package.
- **OfficeIMO:** `OfficeIMO.Core`. PDF parsing, writing, password security, signature structure, logical recovery, manipulation, forms, diagnostics, and preservation analysis are first-party.
- **Optional security:** install `OfficeIMO.Security` for the built-in CMS/X.509/RFC 3161 adapters. It is not a transitive PDF dependency.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
