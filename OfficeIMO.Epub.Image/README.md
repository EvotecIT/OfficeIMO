# OfficeIMO.Epub.Image

`OfficeIMO.Epub.Image` is the thin visual adapter between the bounded EPUB package model and the existing OfficeIMO HTML/Drawing renderer. It adds no image encoder or layout engine.

Retained `application/xhtml+xml` chapters are parsed as XML. Empty non-void elements
remain siblings, and CSS text and SVG/MathML namespaces retain their XML meaning.
Malformed XML, noncanonical XHTML names, processing instructions, and DTD-defined
entity references are rejected instead of being recovered as HTML. `text/html`
chapters and the extracted-text fallback keep HTML parsing.

Load raw chapter HTML and resource payloads when visual fidelity matters:

```csharp
EpubDocument book = EpubDocument.Load(path, new EpubReadOptions {
    IncludeRawHtml = true,
    IncludeResourceData = true
});

IReadOnlyList<OfficeImageExportResult> pages = await book
    .ToImages()
    .Paged()
    .AsPng()
    .ExportAsync();
```

When raw HTML or resource bytes were not retained, export remains available through a diagnosed plain-text or missing-resource fallback.

Package resources use case-sensitive container paths. Chapter count, text-budget,
invalid-encoding, and unsupported-spine omissions propagate into image diagnostics
and the `RequireNoOmissions` policy.

An HTML `base` element resolves relative URLs from the chapter's container path.
External URLs never select retained package bytes. Use asynchronous export with
`ResourceResolver` to supply policy-approved external resources; synchronous
export reports external images that still need resolution.

## Fixed-layout canvas inspection

Inspect one retained XHTML chapter declared `pre-paginated` before exporting it:

```csharp
EpubFixedLayoutInspection inspection = book.InspectFixedLayoutPage(0);
foreach (OfficeDrawingQualityIssue issue in inspection.CanvasQuality.Issues) {
    Console.WriteLine(issue.Message);
}
```

The chapter must declare one viewport with positive numeric `width` and `height`.
Inspection renders at that viewport with zero margins, using the existing package
resource resolver and caller-supplied font/resource limits. It compares rendered
element rectangles with the declared page canvas, including content extending
beyond the automatic output clip. Explicit authored clips remain in effect.
Negative and transformed coordinates are included. Affine effect groups are checked
through their contents; empty space in intermediate rendering buffers does not
count as overflow. Excessive inspection surfaces
fail instead of returning incomplete results. No output image is encoded.

Inspect `HasCanvasOverflow`, `HasTextInkOverflow`, `TextInkDiagnostics`, `ClippingDiagnostics`, `Rendering.Diagnostics`, `PackageDiagnostics` and
`PreparationDiagnostics` together. `HasRenderingWarnings` summarizes diagnosed
rendering loss and package/preparation/inspection warnings. Suppressing package diagnostics for
image export does not suppress them in this inspection. Missing raw XHTML, a
reflowable chapter, an unsupported viewport or encrypted content rejects inspection;
text fallback cannot establish fixed geometry.

Inspect identified layout regions in the same render pass when a page contains
positioned text or image containers:

```csharp
EpubFixedLayoutInspection inspection = book.InspectFixedLayoutRegions(
    0, new[] { "caption", "illustration" });
foreach (EpubFixedLayoutRegionInspection region in inspection.Regions) {
    Console.WriteLine($"{region.ElementId}: box={region.HasOverflow}, ink={region.HasTextInkOverflow}");
}
```

Region findings compare rendered element rectangles with the region's local border
box. Moving or rotating the entire region does not change this local containment;
descendant transforms do. Authored clips inside the region remain in effect, while
ancestor clips do not redefine its local box. The page-level canvas findings still
include the whole scene's transforms and clips. Supply up to 1024 distinct IDs of
positioned, floating, flex or grid containers. Missing or duplicate source IDs and
targets without one rendered region, including hidden or unsupported targets,
reject inspection rather than produce an empty successful result. Region selection
does not modify the retained XHTML or the rendered painting.

Each region also exposes `TextInkDiagnostics` and `HasTextInkOverflow` for positioned
XHTML text in that same local space. A glyph can extend outside its container while
both the layout box and the page-level ink check still fit. Descendant transforms
are included; region and ancestor transforms do not change local containment.
Rectangular and single-convex-contour clips on the region and its descendants constrain the measured contours.
Ancestor clips do not suppress local ink findings. Regional ink warnings contribute
to the enclosing report's `HasRenderingWarnings`.

`HasClippedElementBounds` identifies rendered element rectangles that extend outside
rectangular scene clips. `ClippingDiagnostics` gives the source, clip rectangle,
clipped axes and finding count. These informational findings include intentional
image and background crops; they do not automatically mean the design is wrong.
Descendant transforms and clips are included, and an unclipped axis is not treated
as a boundary. The automatic output clip is excluded. Inspection rejects pages with
more than 1024 scene clips instead of silently truncating the checks.

Path-shaped clips, including rounded overflow boxes, produce
`HtmlRenderClipGeometryNotInspected` warnings. Their precise clipping geometry is
not covered by the rectangular element check. Text-ink inspection below can still
measure supported convex paths; these are separate evidence scopes.

`TextInkDiagnostics` separately compares positioned XHTML glyph outlines with the
page canvas, using the shared drawing engine's font fallback, shaping, color-glyph
layers, synthetic styles and affine transforms. Decorations use conservative stroke
bounds. The measurement uses nominal CSS-pixel geometry at scale 1; it is not a
pixel scan. Empty space in a text frame is excluded. Transparent text is ignored.
Rectangular and single-convex-contour clips intersect the individual measured contours before their bounds
are compared with the page or region. Nested, transformed and single-axis clips
are supported. Rounded rectangles, ellipses and convex polygons use the shared
Drawing renderer's curve flattening at nominal scale 1, including affine transforms.
`HtmlRenderClippedTextInkBounds` informational findings identify
text outline or conservative decoration bounds cropped by a supported authored clip;
review whether the crop is intentional. Automatic output clipping is excluded.
Concave, self-intersecting, multi-contour, degenerate or over-budget paths, embedded vector drawings and unavailable text outlines produce
`HtmlRenderTextInkNotInspected` warnings rather than an apparently clean result.
Supported paths are limited to 512 commands and 512 flattened vertices.
Inspection rejects more than 4096 text runs, 64 nested clips, or
excessive contour-clipping work instead of returning a partial report.

This is managed layout evidence. It does not resolve winding cancellation between overlapping glyph contours, shadows/filters,
or pixel visibility within clipped contours or element rectangles. SVG spine inspection, native-reader presentation and accessible reading
order require separate qualification. A clean report does not certify a publication.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Export | 0 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Epub.Image` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
