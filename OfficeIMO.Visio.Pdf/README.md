# OfficeIMO.Visio.Pdf

`OfficeIMO.Visio.Pdf` converts Visio documents to searchable PDFs. Choose a
semantic report or physical diagram pages. `OfficeIMO.Visio` owns source
projection, and `OfficeIMO.Pdf` owns PDF composition. Reader packages are not
part of this conversion boundary.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

VisioDocument diagram = VisioDocument.Load("architecture.vsdx");
diagram.SaveAsPdf("architecture.pdf");

byte[] pdf = diagram.ToPdfBytes();
```

The default `SemanticReport` mode produces searchable diagram text and topology
through the shared `OfficeDocumentModel`. Its report identifies semantic
fallback rather than claiming native Visio page appearance.

## Diagram pages

`DiagramPages` mode projects the cached artwork into shared Drawing scenes and
creates one PDF page for each source page, in source order. Page dimensions
include the drawing scale; PDF margins are zero. Blank and background pages
remain separate pages. Each page also paints its associated background chain,
deepest background first, with each background's own physical scale and layer
policy. Composition uses the physical lower-left origin without resizing a
background to the foreground page.
Artwork crossing a page edge keeps its original placement; the PDF page bounds
clip the visible result.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

VisioDocument diagram = VisioDocument.LoadLegacyXml("architecture.vdx").Value;
var drawingOptions = new VisioDrawingOptions {
    LayerMode = VisioLayerRenderMode.Printable
};
// Supply licensed font bytes under the source family name when needed.
drawingOptions.Fonts.Add("Arial", File.ReadAllBytes("fonts/Arial.ttf"));
var result = diagram.ToPdfDocumentResult(new VisioToPdfOptions {
    Mode = VisioPdfProjectionMode.DiagramPages,
    DrawingOptions = drawingOptions
});
foreach (var diagnostic in result.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.Location}: {diagnostic.Code}: {diagnostic.Message}");
result.Save("architecture.pdf");
```

The [cached page projection](../OfficeIMO.Visio/README.md#diagram-page-scenes)
retains supported outlines, physical strokes, text runs and paragraphs,
connector routes and bitmap placements. The selected layer policy filters both
painted and searchable page content, including associated backgrounds.
Missing or cyclic background references remain preserved in the source and
report omissions. Native rendering acceptance for background chains with
different page sizes or scales remains open. ShapeSheet recalculation, themes, data
graphics, connector fill artwork, native text fitting and exact
arrow appearance remain outside the qualified profile. Cached shape reflections
follow the [native placement contract](../OfficeIMO.Visio/README.md#cached-shape-text-frame-placement); native reflected-text appearance acceptance remains open. Shared text fitting and curve
flattening report approximations; content that still cannot fit a frame and
unprojected content report omissions.

`DrawingOptions` supplies the same font faces and shaping provider to the scenes
and PDF font resolver. `PdfOptions` configures PDF metadata, security, font
resources and generation policy. Explicit PDF font families take precedence
when overlaid with drawing fonts. Both options are copied for the operation.
Text-clipping diagnostics use the effective PDF font and shaping resources,
including those overrides.
`VisioOptions` and `ProjectionOptions` belong to `SemanticReport`; combining
them with `DiagramPages`, or supplying diagram settings in semantic mode,
throws instead of ignoring settings.

`SourceConversionReports` retains the typed `VisioDrawingConversionReport`.
`RequireNoLoss` on drawing options rejects reported loss before PDF composition;
`SaveLossless` on the PDF result rejects it before writing the destination.
`SaveAsPdfResult` and `SaveAsPdfResultAsync` retain the source report and its
page-qualified fidelity diagnostics when strict projection fails. Their failure
result exposes `HasLoss` and leaves the destination unchanged.
These checks also reject the declared text-layout approximations. Stencil
masters need explicit instantiation on a page before diagram conversion. This
contract does not establish Microsoft Visio open/edit/save acceptance.

## Page previews and legacy XML

Loaded VDX, VSX and VTX documents use the same conversion path. To include a
raster diagram preview alongside searchable semantic content:

```csharp
VisioDocument diagram = VisioDocument.LoadLegacyXml("architecture.vdx").Value;
var conversion = diagram.ToPdfDocumentResult(new VisioToPdfOptions {
    SourceName = "architecture.vdx",
    VisioOptions = new VisioDocumentProjectionOptions {
        IncludePngPreviewAssets = true,
        PngOptions = new VisioPngSaveOptions {
            PixelsPerInch = 96,
            LayerMode = VisioLayerRenderMode.Printable
        }
    }
});
conversion.Save("architecture.pdf");

foreach (var diagnostic in conversion.Warnings) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

PNG previews use visible screen layers by default. The example selects printable layers independently of screen visibility. Semantic text and topology still include hidden content; preview selection does not filter searchable PDF content. See the [layer-selection contract](../OfficeIMO.Visio/README.md#layer-selection-in-previews) for groups, connectors and multiple memberships.

PNG previews render the supported [embedded-image profile](../OfficeIMO.Visio/README.md#legacy-visio-xml),
including crop offsets, master inheritance and group transforms. PDF composition
reduces oversized previews proportionally to fit the configured page content
area. The embedded raster retains its pixels unless the configured image
optimization policy changes them.

Preview diagnostics reach the neutral model, Reader and PDF report with their
source page, preview kind and fidelity category. Unsupported foreign objects
remain visible placeholders and report omission; bitmap header normalization
is informational. Raster allocation limits and cancellation apply during
preview generation. SVG previews remain assets listed as metadata by the PDF
projection; they are not embedded as vector page content.

`SemanticReport` produces a semantic PDF with optional raster previews. It does not reproduce
native Visio page layout, create editable PDF diagram objects, or establish
Microsoft Visio open/edit/save acceptance.

## Dependency footprint

- `OfficeIMO.Core` owns the neutral document model, Drawing scenes, font resources and geometry contracts.
- `OfficeIMO.Visio` owns diagram inspection, preview rendering and projection into those models.
- `OfficeIMO.Pdf` owns loss-aware PDF composition.
- `OfficeIMO.Reader.Visio` and `OfficeIMO.Reader.Pdf` are not dependencies.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Visio.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
