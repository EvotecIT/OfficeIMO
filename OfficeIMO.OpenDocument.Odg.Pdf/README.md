# OfficeIMO.OpenDocument.Odg.Pdf

`OfficeIMO.OpenDocument.Odg.Pdf` converts ODG and FODG drawings to PDF through the
Draw scene projection, the shared Drawing model, and the first-party PDF writer.

```csharp
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;

OdgDocument drawing = OdgDocument.Load("diagram.odg");
// For flat XML, use OdgDocument.LoadFlatXml("diagram.fodg").
var conversion = drawing.ToPdfDocumentResult();
conversion.Save("diagram.pdf").RequireSuccess();

foreach (var diagnostic in conversion.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.Location}: {diagnostic.LossKind}: {diagnostic.Message}");
```

Each source page produces one PDF page, including blank pages, with its original
dimensions and order. Supported text, including image captions, remains searchable. Dynamic page-number/count fields use the current drawing page context; fixed numbers and date/time fields default to reported display snapshots. `OdgToPdfOptions.DateTimeFields` opts into the shared saved-value formatter or dynamic refresh from a supplied timestamp. See the [Draw field profile](../OfficeIMO.OpenDocument/README.md#draw-fields) for formatting and native qualification limits.
Image crop and mirror settings affect the pixels; captions retain the original
frame's text area. PDF output uses
print-visible layers by default; set `OdgToPdfOptions.ForPrint = false` for the
screen-visible view. The converter leaves the source document and its associated
destination unchanged. It uses no external application at runtime.

`ToPdfBytes`, path and caller-owned stream `SaveAsPdf` overloads, asynchronous
save methods, and structured `SaveAsPdfResult` methods use the same engine.
Cancellation is checked between pages and shapes and during PDF saving.
Loading uses `OdfLoadOptions`; `OdgToPdfOptions.MaximumPages` bounds conversion
to 1,000 pages by default. An empty drawing is rejected.

## Fidelity and acceptance

`SourceConversionReports` contains an `OdfConversionReport` with feature locations
such as `page:2:Network/shape:Router`. PDF-stage warnings remain in `Report` and
are refreshed after serialization. `FidelityDiagnostics` and `HasLoss` include
both stages. Use `RequireNoLoss()` after serialization when save-time font or
layout warnings also matter.

The default `LossPolicy` is `ReportOnly`. `ThrowOnSkippedOrUnsupported` accepts
documented approximations but rejects omitted source features before saving;
`ThrowOnAnyLoss` also rejects approximations. Source projection failures do not
publish a PDF through the save methods.
Structured save failures retain the rejected source report in `ConversionReports`,
`FidelityDiagnostics` and `HasLoss`.

Fixed text-box frames and rectangle labels support the [Draw shrinking profile](../OfficeIMO.OpenDocument/README.md#draw-text-fitting). Source clipping checks use the final `PdfOptions` fonts, including named faces, embedded standard fonts and fallback selection, before applying strict loss policy. For fixed frames and line/connector captions, a `text-clipped` mapping identifies omitted content or clipped text paint caused by frame bounds, paragraph margins, padding, tabs or shared layout limits; adjust the frame, margins, padding, line spacing or font size. Shrinking remains an approximation with a six-point largest-font floor.

The [text-box sizing profile](../OfficeIMO.OpenDocument/README.md#draw-text-box-sizing) resolves absolute instance minima and grows horizontal top-aligned boxes in width, height or both axes, using the final PDF font selection. Width grows before height, which is measured with the selected wrapping option at the resolved width. Horizontal ink and decoration insets are reserved before alignment; capped no-wrap text retains its paragraph and text-area anchor. Matching-unit maxima cap the enabled axis; a cap that clips text fails strict conversion. Text and frame paint use the same resolved dimensions while source XML retains its saved geometry. Unspecified growth axes, relative or inconsistent constraints and fitting combined with growth remain unsupported. A grown box can extend beyond its page; growth does not change page dimensions or promise that all text is visible on-page.

The built-in Helvetica, Times and Courier faces measure vertical ink from Adobe's metrics for the caption's encoded glyphs. These fonts are unembedded unless configured otherwise. A PDF reader can substitute different glyph shapes that extend beyond those bounds, so strict conversion does not guarantee unclipped paint in every reader. Supply licensed embedded font data through `PdfOptions` when the output needs predictable font selection. Unavailable glyph bounds and some supplied-font measurements use conservative envelopes; a tight frame can be rejected even when its preview appears to fit. Exact native typography remains outside this qualification.

Conversion uses the [supported Draw scene profile](../OfficeIMO.OpenDocument/README.md#create-and-edit-draw-documents).
Supported master artwork is painted beneath page content, following the same style and visibility contract as `ToDrawing`.
Literal enhanced paths and full ellipses follow the [enhanced geometry profile](../OfficeIMO.OpenDocument/README.md#draw-enhanced-geometry).
Backgrounds outside the [solid, native gradient and stretched bitmap profile](../OfficeIMO.OpenDocument/README.md#draw-page-backgrounds), enhanced geometry outside that profile, embedded objects, advanced styles and effects,
source metadata and hyperlink targets outside the supported URI profile are not reproduced. Basic geometry
is still classified as an approximation. Native font metrics, arbitrary typography
and complete document fidelity are not established by page or text checks.
Windows .NET Framework runtime and broader platform qualification remain separate
acceptance work.

The `odg-pdf` route in `OfficeIMO.Workflows` accepts both `.odg` and `.fodg` and
returns source and PDF evidence with the output. Its preview and batch consumers
use the same converter.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument.Odg.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
