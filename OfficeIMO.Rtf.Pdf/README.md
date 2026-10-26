# OfficeIMO.Rtf.Pdf

First-party semantic RTF/PDF conversion for OfficeIMO.

This package converts between the semantic `OfficeIMO.Rtf` document model and the first-party `OfficeIMO.Pdf` document model. The RTF engine remains the lossless parse/edit/write layer; PDF export is a visual/content conversion to a fixed-layout format, while PDF import uses the first-party logical PDF reader to recover parser-supported metadata, pages, headings, grouped paragraphs, and list markers into an editable RTF document.

Supported export coverage includes semantic paragraphs, paragraph indentation/spacing/line-height/pagination controls, paragraph/style tab stops with supported alignment and leader mapping, section-owned blocks, section page breaks, page-starting section page setup, document and section page-border visual export, rich runs, list markers, document page setup, metadata, tables with horizontal and vertical merged-cell spans, repeating header rows, solid row/cell fills, cell padding, vertical alignment, side and diagonal cell borders, PNG/JPEG images, bookmarks, field result text, hidden text control, footnote/endnote/annotation bodies, and rich running header/footer paragraphs including first-page and even-page variants. RTF can model separate borders per page side; PDF export maps the first styled RTF page border to the first-party PDF engine's uniform page border decoration.

Supported import coverage includes PDF Info metadata, first-page paper size, logical headings, logical list items, grouped paragraphs, basic paragraph spacing, and page transitions as RTF page-break-before paragraphs. PDF is a fixed-layout format, so import is semantic text extraction rather than lossless visual reconstruction of arbitrary PDFs.

Page-starting sections inherit document page defaults and apply their own size, margins, orientation, page numbering, and header/footer declarations. First-page stories require `DifferentFirstPageHeaderFooter`; even-page stories require `Settings.FacingPages`. Odd/even selection follows the visible page number, while the first-page variant belongs to each section. Odd/even section breaks and facing-page numbering restarts insert a blank page when needed to align the physical pages. Document page-number starts continue across sections unless a section requests a restart.

Equal-width section columns flow through the shared PDF layout engine, including column gaps, separator lines, and explicit column and page breaks. Continuous sections can switch between single-column and equal-width column flow on the current page. Unequal or right-to-left columns produce `SectionColumnsFlattened`; changes to page geometry, numbering, or running stories inside a continuous section produce `ContinuousSectionPageSettingsFlattened`. Column-starting section breaks use a new page and produce `ColumnSectionBreakFlattened` when multiple columns are involved. The report identifies these limits while retaining the body content.

## Paragraph layout

Pair kerning is disabled unless a run supplies a positive `KerningHalfPoints` threshold. The threshold is expressed in half-points; kerning applies when the effective font size meets or exceeds it. The size includes inherited RTF formatting or the PDF default when no run size resolves. Source `\kerning0` disables kerning.

PDF export distinguishes minimum, exact, and multiple line spacing using the [RTF sign and unit rules](../OfficeIMO.Rtf/README.md#paragraph-spacing-and-inheritance). Body paragraphs, individual table-cell paragraphs, and appended note paragraphs retain supported alignment, indents, before/after spacing, line spacing, and tab stops. An ordinary RTF paragraph without after-spacing does not acquire the PDF builder's convenience gap.

Bar tab stops draw vertical rules at their declared positions without moving the text cursor, including paragraphs with no tab character and supported rich running stories. The shared PDF engine supplies their default appearance; `BarTabAppearanceApproximated` records `Approximation` loss because native RTF pen widths and line extents are not independently qualified. Bar rules in nested tables or caller-selected plain running text are omitted and reported as `BarTabOmitted` with `Omission` loss. `RequireNoLoss()` rejects both diagnostics.

RTF export defaults to `PdfTextWhitespaceMode.Preserve`: it retains repeated source spaces and moves a whole hyphenated token to the next line when it fits there. An explicit `TextWhitespaceMode` in `RtfToPdfOptions.PdfOptions` overrides that default. Use `Collapse` for collapsed spaces or `Preformatted` for literal-space overflow behavior. Conversion leaves the caller's options unchanged.

By default, borderless zero-gap RTF tables do not acquire a PDF grid or convenience cell padding. A nonnegative row half-gap supplies left/right cell padding when no explicit cell or row padding applies. Explicit cell padding, including zero, overrides row padding and the half-gap. A caller-selected `PdfOptions.DefaultTableStyle` supplies fallback padding instead of the half-gap; explicit source padding and borders still apply.

When the selected font programs and resolved run sizes provide usable line metrics, automatic, minimum, and multiple spacing use those metrics. Otherwise, the shared PDF spacing policy remains in effect. Exact spacing keeps the requested advance. Font substitution can change line heights and page breaks; this supported metric profile does not establish identical native Word layout for arbitrary documents.

Footnotes, endnotes, and annotations are appended after the body with a label on each note's first paragraph. Their paragraph formatting is retained within that flow; this route does not place footnotes beside their reference pages or reproduce native endnote continuation. Nested tables remain delimited text with `NestedTableFlattened` diagnostics.

Paragraph-frame positioning, dimensions, and wrapping are flattened to ordinary flow and reported as `ParagraphFrameFlattened`. The warning applies to projected body, cell, note, and selected running-story paragraphs. `RequireNoLoss()` rejects this approximation; the lossless RTF API remains available when source frame syntax must survive.

## Running headers and footers

Selected running stories retain paragraph alignment, line spacing, tabs, supported run formatting and inline images in source order. Header and footer distances are measured from the page edge; an absent distance uses the RTF half-inch default. Running-story fonts share the document's existing font policy and resolution budget. Missing declarations inherit, while an explicitly empty declaration clears its selected story.

Ordinary inline RTF pictures retain their full authored frame on exact-height running lines. The shared PDF engine measures the story and reserves its body space. First-page and even-page selection, numbering restarts and intentional facing-page blanks retain their section semantics; a blank physical page is not automatically an error.

Explicit header/footer font, alignment, color or offset presentation settings select the existing plain-text projection for that side. Images omitted by this projection report `ImageSkipped`. `IncludeHeaderFooters = false` skips source stories. `IncludeImages = false` uses the existing image omission policy within supported rich stories.

Rich running stories are PDF pagination artifacts. Read them with `PdfReadDocument.Open(bytes, new PdfLoadOptions { IncludeArtifactText = true })` when inspecting header/footer text. The logical reader otherwise excludes artifact text by default.

In continuous, single-column paragraph flow, a change between existing ordinary default headers retains the current physical page's header and selects the resolved replacement on the next physical page. The new header uses its source fonts and reserves its own measured body space. This profile requires unchanged page geometry, running-story distances, page numbering and footer, without active line numbering, first/even variants, notes, tables, paragraph frames or complex body/story content. Other continuous-story changes retain the reported approximation path.

An empty body section retains its implicit RTF paragraph in both direct model conversion and saved/reloaded input. Rendering that paragraph leaves the source section's block collection unchanged.

Running content is bounded: page and column controls cannot create new pages, note bodies are not appended from a running story, and fields use cached result text rather than per-page evaluation. Unsupported controls and source formatting retain operation diagnostics.

## Import PDF with diagnostics

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Rtf.Pdf;

PdfDocument pdf = PdfDocument.Load("source.pdf");
PdfRtfConversionResult result = pdf.ToRtfDocumentResult(new PdfToRtfOptions {
    ReadOptions = new PdfReadOptions {
        PageSelection = PdfPageSelection.Parse("1-10"),
        Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 10 }
    }
});

result.Value.Save("source.rtf");
foreach (var warning in result.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

`ReadOptions` is the same canonical semantic configuration used by `PdfDocument.Read(...)`, including page selection, Fast/Structured profile choice, layout settings, and resource budgets. It is ignored when converting an existing `PdfDocumentReadResult`.

The semantic importer preserves detected run color, font size, bold, and italic through shared logical PDF runs. Tables, images, links, and form widgets currently produce structured loss diagnostics instead of disappearing silently.

## Export with diagnostics

```csharp
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Pdf;

RtfDocument rtf = RtfDocument.Load("input.rtf");
var options = new RtfToPdfOptions();
PdfDocumentConversionResult result = rtf.ToPdfDocumentResult(options);

result.Report.RequireNoErrorWarnings();
result.Save("output.pdf");
```

`RtfToPdfOptions.TextFallbacks` defaults to `PdfTextFallbackFeatures.Default`, including installed multilingual fonts. Fallback faces supply characters missing from authored fonts. Set it to `None` for exact document-font output.

All PDF adapters expose the same `PdfResourcePolicy`. RTF conversion uses the balanced default: it may discover and embed installed fonts, but it does not allow arbitrary local-file or remote-resource reads. Use `PdfResourcePolicy.CreatePortableDeterministic()` when conversion must not inspect host fonts. RTF reads document content only from the typed model, so `CreateTrustedHost()` adds no access unless a format-specific resolver is explicitly configured.

For raw RTF strings, bytes, or streams, use source-explicit APIs such as `ToPdfFromRtf()`, `ToPdfDocumentFromRtf()`, and `SaveAsPdfFromRtf()`. Typed `RtfDocument` instances use the standard `ToPdfBytes()`, `ToPdfDocument()`, and destination-only `SaveAsPdf()` names.

PNG, JPEG, and supported DIB images use the shared managed drawing layer. Set `RtfToPdfOptions.ImageConverter` when WMF/EMF content must be rasterized; a null or invalid callback result is reported rather than silently treated as an image.

## Dependency footprint

- **External:** None by default. Applications may provide an optional WMF/EMF rasterization callback.
- **OfficeIMO:** `OfficeIMO.Rtf`, `OfficeIMO.Pdf`, and `OfficeIMO.Core` own semantic mapping, rendering, logical import, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 1 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Rtf.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
