using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Visio.Tests;

public sealed class VisioDrawingPdfFidelityTests {
    [Fact]
    public void DefaultPdfKeepsNativeCurvesAndCachedGradientFillsAsVectorPaint() {
        string fixture = Path.Combine(OfficeIMO.Tests.RepositoryTestPaths.Find(), "Assets", "VisioTemplates", "DrawingWithJenkinsDiagram.vsdx");
        VisioDocument document = VisioDocument.Load(fixture);

        PdfDocument conversion = document.ToPdf(DiagramOptions(), out OfficeConversionReport report);
        PdfDocument pdf = PdfDocument.Load(new MemoryStream(conversion.ToBytes()));
        List<OfficeShape> shapes = Flatten(Assert.Single(pdf.Pages).ToDrawing()).ToList();

        // Ellipse and elliptical-arc rows reach the PDF as Bezier segments rather than polylines.
        Assert.Contains(shapes, shape => shape.PathCommands.Any(command => command.Kind == OfficePathCommandKind.CubicBezierTo));
        // The Azure icons' cached FillGradient stops are painted as a shading, not their foreground color.
        Assert.Contains(shapes, shape => shape.FillGradient?.Stops.Any(stop => stop.Color.R == 0x00 && stop.Color.G == 0x78 && stop.Color.B == 0xD4) == true);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "VISIO_DRAWING_CURVES");
    }

    [Fact]
    public void ArrowKindsKeepDistinctShapesAndHyperlinksBecomePdfLinks() {
        VisioDocument document = VisioDocument.Create(); document.UseMastersByDefault = false;
        VisioPage overview = document.AddPage("Overview", OfficeLength.Inches(6), OfficeLength.Inches(3));
        VisioShape source = overview.AddRectangle(OfficeLength.Inches(1), OfficeLength.Inches(1.5), OfficeLength.Inches(1), OfficeLength.Inches(0.5), "Source");
        VisioShape target = overview.AddRectangle(OfficeLength.Inches(5), OfficeLength.Inches(1.5), OfficeLength.Inches(1), OfficeLength.Inches(0.5), "Target");
        source.AddHyperlink("https://example.com/docs", "Documentation");
        target.AddPageHyperlink("Details");
        VisioConnector connector = overview.AddConnector("flow", source, target, ConnectorKind.Straight);
        connector.BeginArrow = (EndArrow)20;
        connector.EndArrow = (EndArrow)1;
        document.AddPage("Details", OfficeLength.Inches(6), OfficeLength.Inches(3))
            .AddRectangle(OfficeLength.Inches(3), OfficeLength.Inches(1.5), OfficeLength.Inches(1), OfficeLength.Inches(0.5), "Details");

        OfficeDrawing scene = document.ToDrawings().Value[0];
        List<OfficeShape> shapes = Flatten(scene).ToList();
        // Kind 1 is an open line arrowhead: stroked, unfilled and not closed.
        Assert.Contains(shapes, shape => shape.FillColor == null && shape.StrokeColor != null
            && shape.PathCommands.Count == 3 && shape.PathCommands.All(command => command.Kind != OfficePathCommandKind.Close));
        // Kind 20 is a filled dot, not the triangle every kind used to share.
        Assert.Contains(shapes, shape => shape.FillColor != null && shape.StrokeColor == null && shape.PathCommands.Count > 10);

        PdfDocument conversion = document.ToPdf(DiagramOptions(), out OfficeConversionReport report);
        PdfDocument pdf = PdfDocument.Load(new MemoryStream(conversion.ToBytes()));
        IReadOnlyList<PdfLinkAnnotation> links = pdf.Pages[0].GetLinkAnnotations();
        Assert.Contains(links, link => link.Uri == "https://example.com/docs");
        Assert.Contains(links, link => link.DestinationName == "visio-page-2");
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code is "VISIO_DRAWING_METADATA" or "VISIO_DRAWING_LINKS" or "VISIO_DRAWING_ARROW_KIND");
    }

    [Fact]
    public void SvgPreviewPaintsTheSameCachedGradientStopsAsTheDiagramScene() {
        VisioDocument document = VisioDocument.Load(Path.Combine(OfficeIMO.Tests.RepositoryTestPaths.Find(), "Assets", "VisioTemplates", "DrawingWithJenkinsDiagram.vsdx"));

        string svg = document.ToSvg();

        // The master-inherited stops reach the preview as a gradient paint, not the black foreground color.
        Assert.Contains("<linearGradient", svg, StringComparison.Ordinal);
        Assert.Contains("0078d4", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("fill=\"url(#visio-fill-", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void TextBackgroundFromTheTextStyleSheetIsPaintedAndAnUnsetOneStaysTransparent() {
        string root = OfficeIMO.Tests.RepositoryTestPaths.Find();
        // The Jenkins dynamic connector master uses the Connector style sheet, whose TextBkgnd is white.
        VisioDocument styled = VisioDocument.Load(Path.Combine(root, "Assets", "VisioTemplates", "DrawingWithJenkinsDiagram.vsdx"));
        VisioConnector connector = styled.Pages[0].Connectors.First(item => VisioConnectorGeometry.HasVisibleLine(item));
        connector.Label = "Styled label";
        // This document defines no TextBkgnd anywhere, so Visio's transparent default applies.
        VisioDocument unset = VisioDocument.Load(Path.Combine(root, "Website", "static", "downloads", "showcase", "visio", "premium-cloud-architecture.vsdx"));

        Assert.Contains(LabelBackings(styled.ToDrawings().Value[0], "Styled label"), color => color.R == 255 && color.G == 255 && color.B == 255 && color.A == 255);
        Assert.Empty(LabelBackings(unset.ToDrawings().Value[0], "auth"));
    }

    private static IEnumerable<OfficeColor> LabelBackings(OfficeDrawing scene, string label) =>
        scene.Elements.OfType<OfficeDrawingEffectGroup>()
            .Select(group => group.Drawing)
            .Where(local => local.Elements.OfType<OfficeDrawingRichText>().Any(text => text.PlainText.Contains(label)))
            .SelectMany(local => local.Shapes.Select(item => item.Shape.FillColor).OfType<OfficeColor>());

    private static VisioPdfExportOptions DiagramOptions() => new() { Mode = VisioPdfProjectionMode.DiagramPages };

    private static IEnumerable<OfficeShape> Flatten(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            IEnumerable<OfficeShape> nested = element switch {
                OfficeDrawingShape shape => new[] { shape.Shape },
                OfficeDrawingEffectGroup group => Flatten(group.Drawing),
                OfficeDrawingGroup group => Flatten(group.Drawing),
                _ => Enumerable.Empty<OfficeShape>()
            };
            foreach (OfficeShape shape in nested) yield return shape;
        }
    }
}
