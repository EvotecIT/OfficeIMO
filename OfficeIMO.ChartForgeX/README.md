# OfficeIMO.ChartForgeX

`OfficeIMO.ChartForgeX` is the optional bridge for placing any ChartForgeX `VisualArtifact` in Word, Excel, PowerPoint, PDF, or another `OfficeDrawing` consumer, and for projecting supported diagram semantics into native editable Visio. Existing OfficeIMO packages do not acquire a ChartForgeX dependency.

For source builds, reference `OfficeIMO.ChartForgeX.csproj` from the consuming project and make the ChartForgeX source projects available through the repository's project-reference configuration. The bridge remains optional: applications that do not reference it keep the standard OfficeIMO dependency graph.

Every CFX surface that emits SVG can use the flat Office placement path, even when it does not expose a typed artifact envelope. Wrap the generated markup in `OfficeVisualSource`:

```csharp
OfficeVisualConversionResult visual = new OfficeVisualSource(canvas.ToSvg()) {
    Id = "release-overview",
    Title = "Release overview",
    AlternativeText = "Release readiness summary with six status tiles."
}.ToOfficeVisual();
```

The bridge renders once and returns an `OfficeVisualConversionResult` containing:

- SVG or PNG placement bytes for Word, Excel, and PowerPoint, according to the selected SVG policy;
- an `OfficeDrawing` scene for PDF and drawing pipelines;
- dimensions normalized to points;
- accessible text, metadata-ready regions, and a typed fidelity report.

```csharp
using ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;

VisualArtifact artifact = chart.ToVisualArtifact("sales-quarter");
artifact.Accessibility.WithTextAlternative(
    "Quarterly sales",
    "Revenue increased in each of the four reported quarters.");

OfficeVisualConversionResult visual = artifact.ToOfficeVisual(
    new OfficeVisualConversionOptions {
        WidthPoints = 420,
        SvgPolicy = OfficeVisualSvgPolicy.RasterizeWhenNeeded
    });

paragraph.AddVisualArtifact(visual);
sheet.AddVisualArtifact(2, 2, visual);
slide.AddVisualArtifact(visual, leftPoints: 36, topPoints: 72);

PdfDocument.Create(pdf => pdf.Content(content =>
    content.AddVisualArtifact(visual)), pdfOptions);
```

The default `PreserveVector` policy keeps the imported vector scene and reports SVG features that OfficeIMO.Drawing cannot represent. Choose `RasterizeWhenNeeded` for visual fidelity when unsupported SVG features should use the PNG placement payload, or `RequireVector` when incomplete vector conversion must fail closed. Word, Excel, and PowerPoint use the selected placement payload; PDF uses the converted `OfficeDrawing` scene.

## Native editable Visio

Topology, flow, and sequence artifacts can be projected into native OfficeIMO.Visio diagrams. Nodes, containers, connectors, Shape Data, hyperlinks, sequence messages, activations, notes, and fragments remain editable after saving to VSDX. The conversion result includes the document, generated page, validated CFX interchange envelope, and a fidelity report.

```csharp
using ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;

VisualArtifact artifact = topology.ToVisualArtifact();
OfficeVisioVisualConversionResult visio = artifact.ToOfficeVisio(
    new OfficeVisioVisualOptions { PageName = "Service topology" });

visio.Document.Save("service-topology.vsdx");
```

Use `artifact.ToInterchangeUtf8Json()` and `jsonBytes.ToOfficeVisio()` across process or PowerShell assembly-load-context boundaries. Static SVG remains a separate fallback; the adapter does not infer editable semantics by scraping rendered markup. Unsupported artifact families fail closed for native Visio conversion. Native builders fit the page to editable content by default; set `UseNaturalPageSize` when the CFX pixel viewport must remain the minimum page size.

ChartForgeX owns chart and diagram semantics, deterministic rendering, interchange, watermarks, layout, and raster metadata. OfficeIMO owns document placement, native Visio projection, page layout, document/page watermarks, PDF composition, and Office package behavior.
