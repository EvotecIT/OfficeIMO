# OfficeIMO.Epub.Image

`OfficeIMO.Epub.Image` is the thin visual adapter between the bounded EPUB package model and the existing OfficeIMO HTML/Drawing renderer. It adds no image encoder or layout engine.

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
Negative and transformed coordinates are included; excessive inspection surfaces
fail instead of returning incomplete results. No output image is encoded.

Inspect `HasCanvasOverflow`, `Rendering.Diagnostics`, `PackageDiagnostics` and
`PreparationDiagnostics` together. `HasRenderingWarnings` summarizes diagnosed
rendering loss and package/preparation warnings. Suppressing package diagnostics for
image export does not suppress them in this inspection. Missing raw XHTML, a
reflowable chapter, an unsupported viewport or encrypted content rejects inspection;
text fallback cannot establish fixed geometry.

This is managed layout evidence. It does not measure glyph ink, shadows/filters,
content hidden by authored clipping, or overflow inside individual positioned
regions. SVG spine inspection, native-reader presentation and accessible reading
order require separate qualification. A clean report does not certify a publication.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Export | 0 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Epub.Image` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
