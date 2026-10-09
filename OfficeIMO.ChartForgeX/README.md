# OfficeIMO.ChartForgeX

`OfficeIMO.ChartForgeX` is the optional bridge for placing any ChartForgeX `VisualArtifact` in Word, Excel, PowerPoint, PDF, or another `OfficeDrawing` consumer, and for projecting supported diagram semantics into native editable Visio. Existing OfficeIMO packages do not acquire a ChartForgeX dependency.

The bridge uses ChartForgeX 2.0. Add `ChartForgeX.Visuals` when producing factual tables, canvases, or watermark decoration. Add `ChartForgeX.Stories` only in applications that produce stories or animation. Both optional producers can supply a common static artifact to the bridge; the bridge itself keeps a core-only ChartForgeX dependency.

For source builds, reference `OfficeIMO.ChartForgeX.csproj` from the consuming project and make the ChartForgeX source projects available through the repository's project-reference configuration. The bridge remains optional: applications that do not reference it keep the standard OfficeIMO dependency graph.

## Put a chart in Word

Prepare the chart at its intended document size, then insert its artifact. `OfficeVisualDocumentStyle` uses point-based typography and the shared light/dark palette. Word constrains an oversized visual to the paragraph's available content width when no explicit size is supplied.

```csharp
using ChartForgeX.Core;
using ChartForgeX.Primitives;
using ChartForgeX.Rendering;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Word;

var chart = Chart.Create().AddLine("Revenue", new[] {
    new ChartPoint(1, 42), new ChartPoint(2, 58), new ChartPoint(3, 73)
});
var context = OfficeVisualDocumentStyle.Default.CreateContext(
    frame: new VisualFrame("Quarterly sales", showLegend: false));
var artifact = chart.Prepare(context).ToArtifact("sales-quarter");
artifact.Accessibility.WithTextAlternative("Quarterly sales", "Revenue: 42, 58 and 73.");

using var document = WordDocument.Create("sales.docx");
document.AddParagraph().AddVisualArtifact(artifact);
document.Save();
```

## Choose the destination

Use the same artifact in each format. Worksheet ranges and slide layout boxes preserve proportions and center the visual within the authored bounds.

```csharp
paragraph.AddVisualArtifact(artifact);
sheet.AddVisualArtifact("B4:I18", artifact);
slide.AddVisualArtifact(artifact, presentation.SlideSize.GetContentBoxPoints(36));
content.AddVisualArtifact(artifact, spacingAfter: 12);
```

Word uses the containing cell or owning section, authored columns, margins and direct paragraph indents. `paragraph.GetContentWidthPoints()` exposes that estimate for preparation. Unequal flowing columns use their smallest authored width. Shared headers with differing section widths and unsupported text boxes require explicit dimensions. This is authored geometry, not Microsoft Word's measured pagination or automatic table layout.

PDF constrains oversized drawings to their actual flow width, including padded containers and columns, without enlarging smaller visuals. Supply a `PdfDrawingStyle` to explicitly choose placement behavior; set `ConstrainToContentWidth = true` to retain automatic fitting with custom styling.

Set `WidthPoints` or `HeightPoints` for an explicit size. With both supplied, the default `Fit = OfficeImageFit.Contain` keeps the whole visual within the box without distortion. Choose `Stretch` explicitly for exact dimensions. Cropped `Cover` fitting is not supported by this bridge. Slide boxes and worksheet ranges supply their own final placement bounds.

## Reuse a conversion and inspect fidelity

The bridge returns an `OfficeVisualConversionResult` containing:

- SVG or PNG placement bytes for Word, Excel, and PowerPoint, according to the selected SVG policy;
- an `OfficeDrawing` scene for PDF and drawing pipelines;
- dimensions normalized to points;
- accessible text, metadata-ready regions, and a typed fidelity report.

```csharp
OfficeVisualConversionResult visual = artifact.ToOfficeVisual(
    new OfficeVisualConversionOptions { WidthPoints = 420 });

paragraph.AddVisualArtifact(visual);
sheet.AddVisualArtifact(2, 2, visual);
slide.AddVisualArtifact(visual, leftPoints: 36, topPoints: 72);

content.AddVisualArtifact(visual);

// Change placement size without rendering or importing the source again.
slide.AddVisualArtifact(visual.WithSize(300), leftPoints: 36, topPoints: 72);
foreach (var diagnostic in visual.Report.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

Insertion overloads with `out OfficeVisualConversionResult` expose the conversion when inserting an artifact directly. Use that result to inspect diagnostics. `Report.RequireNoLoss()` rejects reported approximations and omissions.

The default `RasterizeWhenNeeded` policy preserves appearance with vector output when supported and PNG fallback when SVG import is incomplete. Fallback remains a reported approximation, with loss of vector editability. Choose `PreserveVector` to keep the imported vector scene and report unsupported features, or `RequireVector` to reject incomplete vector conversion. Word, Excel, and PowerPoint use the selected placement payload; PDF uses the converted `OfficeDrawing` scene.

Prepared artifacts retain their resolved viewport and accessible text. The default scale is 0.75 points per pixel; `PointsPerPixel` changes that scale, while `WidthPoints` or `HeightPoints` can set the document size explicitly. Raster DPI metadata does not change the chosen placement size.

Choose a viewport and typography for the intended document size before preparing the visual. Reducing a wide chart to a small picture also reduces every label. `OfficeVisualDocumentStyle.Default.CreateContext(widthPoints, heightPoints)` supplies readable document typography at that size; its font and point sizes are also available for surrounding headings and captions. A custom style can change that scale, while a supplied `VisualTheme` retains the shared palette and geometry. The [document delivery example](../OfficeIMO.ChartForgeX.Examples/README.md) demonstrates light/dark charts, repeated placements, saved-file previews and editable topology fidelity across the five document formats.

## SVG-producing surfaces

Every CFX surface that emits SVG can use the flat Office placement path, even when it does not expose a typed artifact envelope. Wrap the generated markup in `OfficeVisualSource`:

```csharp
OfficeVisualConversionResult visual = new OfficeVisualSource(canvas.ToSvg()) {
    Id = "release-overview",
    Title = "Release overview",
    AlternativeText = "Release readiness summary with six status tiles."
}.ToOfficeVisual();
```

Apply static watermarks before conversion with the Visuals decorator:

```csharp
artifact.WithWatermarks(VisualWatermark.FromText("Internal"));
OfficeVisualConversionResult visual = artifact.ToOfficeVisual();
```

Decoration retains the semantic model. Native Visio reports `WatermarkNotProjected` because its editable page does not project the static watermark layers; the SVG or PNG picture contains them. The same diagnostic and fidelity policy apply when the artifact is passed as an interchange envelope or UTF-8 JSON.

## Native editable Visio

Topology, flow, and sequence artifacts can be projected into native OfficeIMO.Visio diagrams. Nodes, containers, connectors, Shape Data, hyperlinks, sequence messages, activations, notes, and fragments remain editable after saving to VSDX. The conversion result includes the document, generated page, validated CFX interchange envelope, and a fidelity report.

```csharp
using ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;

VisualArtifact artifact = topology.Prepare().ToArtifact("service-topology");
OfficeVisioVisualConversionResult visio = artifact.ToOfficeVisio(
    new OfficeVisioVisualOptions { PageName = "Service topology" });

visio.Document.Save("service-topology.vsdx");

if (visio.Report.HasSemanticLoss) {
    foreach (OfficeVisioVisualDiagnostic diagnostic in visio.Report.Diagnostics) {
        Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
    }
}
```

`AllProjectedObjectsEditable` reports whether the native objects can be edited independently. `HasSemanticLoss` is a separate fidelity decision: it becomes true only for warning diagnostics, when Visio cannot represent a source feature exactly, even though every projected shape remains editable. Disabling Shape Data or native hyperlinks is therefore reported when source semantics would otherwise be omitted. Information diagnostics describe lossless normalizations such as collision-safe Shape Data aliases. `ArtifactKind` preserves how the visual was authored, while `SemanticFamily` identifies the reusable topology, flow, or sequence projection. Use stable diagnostic codes and severity for automation and `Message` for logs or user interfaces.

Pass a typed `VisualArtifactInterchangeEnvelope` directly when the caller shares the same ChartForgeX assembly identity. The adapter validates that object without serializing it first. Use `artifact.ToInterchangeUtf8Json()` and `jsonBytes.ToOfficeVisio()` across process or PowerShell assembly-load-context boundaries. Static SVG remains a separate fallback; the adapter does not infer editable semantics by scraping rendered markup. Unsupported artifact families fail closed for native Visio conversion. Complete topology envelopes preserve node and group bounds by default. Coordinates use `PixelsPerInch` to convert CFX pixels to Visio inches. Unpositioned topology, flow, and sequence inputs use native layout; `UseNaturalPageSize` keeps the CFX viewport as their minimum page size.

ChartForgeX owns chart and diagram semantics, deterministic rendering, interchange, watermarks, layout, and raster metadata. OfficeIMO owns document placement, native Visio projection, page layout, document/page watermarks, PDF composition, and Office package behavior.

### Preserve placement and enforce fidelity

`LayoutMode = Auto` preserves complete topology bounds and reflows other inputs. Choose `Preserve` to require prepared topology coordinates, or `Reflow` to explicitly request native layout. Prepared topology graphs retain their resolved connector routes and label rectangles. Envelopes without a resolved route retain authored connector bends and named port offsets; native route construction is reported when neither is available. Native titles use clear space above the preserved content; when no header band is available, the title is omitted and `TitleNotProjected` is reported. Native graph styling maps the source page background, card, surface, border, and foreground colors, with Arial text for portable previews. The page fill is a protected native adornment behind the editable graph objects; it keeps the same page count and gives titles and connector captions the intended light or dark surface. Set `NativeTheme` to override it, for example with `VisioStyleTheme.Technical()` for the previous native defaults. Curves, advanced edge styling, icons, source fonts, and complete CFX themes still have limits described by the diagnostics.

Flow and sequence remain native editable diagrams with recomputed layout. Their prepared coordinates remain in the interchange envelope, and `LayoutRecomputed` reports that the native page uses a different layout. `Preserve` rejects those families instead of claiming exact placement.

```csharp
var options = new OfficeVisioVisualOptions {
    LayoutMode = OfficeVisioVisualLayoutMode.Preserve,
    PixelsPerInch = 96
};
options.RejectedDiagnostics.Add(OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
var result = envelope.ToOfficeVisio(options);
result.Document.Save("topology.vsdx");
```

Set `RequireLossless = true` to reject every warning, or add selected codes to `RejectedDiagnostics`. A rejected conversion throws `OfficeVisioVisualFidelityException`; its `Report` explains the losses. Conversion does not save a file.

### Export a report as one document

Pass an ordered sequence of envelopes to `ToOfficeVisioBook`. Each input becomes one page; repeated titles receive distinct page names. Each result in `Pages` exposes that page's fidelity report and shares the returned document.

```csharp
var book = envelopes.ToOfficeVisioBook(options);
book.Document.Save("topology-report.vsdx");
```

ChartForgeX report pages can supply these envelopes through `report.Pages.Select(page => page.ToInterchangeEnvelope())`. Project `CrossPageLinks` into native page links to make every relationship navigable in Visio:

```csharp
var envelopes = report.Pages.Select(page => page.ToInterchangeEnvelope());
var links = report.CrossPageLinks.Select(link => new OfficeVisioVisualBookLink(
    link.SourcePage, link.SourceNodeId,
    link.TargetPage, link.TargetNodeId,
    link.EdgeId));

var book = envelopes.ToOfficeVisioBookWithNavigation(links, new OfficeVisioVisualBookOptions {
    MaximumNavigationLinksPerEntity = 12,
    IncludeReturnLinks = true
});
book.Document.Save("topology-report.vsdx");
```

Duplicate targets are coalesced while their relationship identifiers remain in Shape Data. `OmittedNavigationCount` and `CFX.BookLink.Omitted` disclose links removed by the per-entity bound. ChartForgeX retains an original node identifier in `chartforgex.sourceId` whenever interchange projection has to shorten or disambiguate it, so report links still resolve against the editable Visio shape.
