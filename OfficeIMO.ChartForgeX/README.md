# OfficeIMO.ChartForgeX

`OfficeIMO.ChartForgeX` is the optional bridge for placing any ChartForgeX `VisualArtifact` in Word, Excel, PowerPoint, PDF, or another `OfficeDrawing` consumer. Existing OfficeIMO packages do not acquire a ChartForgeX dependency.

Package publication is intentionally deferred while the ChartForgeX 1.5 and OfficeIMO integration APIs stabilize. Until the coordinated packages are published, build this repository with the adjacent ChartForgeX source checkout or use a project reference; do not expect `dotnet add package OfficeIMO.ChartForgeX` to resolve from NuGet yet.

Every CFX surface that emits SVG can use the same bridge, even when it does not expose a typed artifact envelope. Wrap the generated markup in `OfficeVisualSource`; this is also the stable exchange contract across processes and isolated PowerShell module load contexts:

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

ChartForgeX owns chart and diagram rendering, watermarks, layout, and raster metadata. OfficeIMO owns document placement, page layout, document/page watermarks, PDF composition, and Office package behavior.
