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
    public void ConnectorLabelBackgroundFollowsTheFileInPdfSvgAndPng() {
        string root = OfficeIMO.Tests.RepositoryTestPaths.Find();
        // The Jenkins dynamic connector master uses the Connector style sheet, whose TextBkgnd is white.
        VisioDocument styled = VisioDocument.Load(Path.Combine(root, "Assets", "VisioTemplates", "DrawingWithJenkinsDiagram.vsdx"));
        VisioConnector connector = styled.Pages[0].Connectors.First(item => VisioConnectorGeometry.HasVisibleLine(item));
        connector.Label = "Styled label";
        // This document defines no TextBkgnd anywhere, so Visio's transparent default applies.
        VisioDocument unset = VisioDocument.Load(Path.Combine(root, "Website", "static", "downloads", "showcase", "visio", "premium-cloud-architecture.vsdx"));

        Assert.Contains(LabelBackings(styled.ToDrawings().Value[0], "Styled label"), color => color.R == 255 && color.G == 255 && color.B == 255 && color.A == 255);
        Assert.Empty(LabelBackings(unset.ToDrawings().Value[0], "auth"));

        System.Xml.Linq.XElement styledLabel = SvgConnector(styled.Pages[0], connector.Id);
        Assert.Contains(styledLabel.Elements(), element => (string?)element.Attribute("data-officeimo-connector-label-background") == "true"
            && string.Equals((string?)element.Attribute("fill"), "#FFFFFF", StringComparison.OrdinalIgnoreCase) && element.Attribute("fill-opacity") == null);
        Assert.DoesNotContain("data-officeimo-connector-label-background", unset.Pages[0].ToSvg(), StringComparison.Ordinal);

        Assert.True(LinePixelsCoveredByLabelBackings(styled.Pages[0]) > 0);
        Assert.Equal(0, LinePixelsCoveredByLabelBackings(unset.Pages[0]));
    }

    private static System.Xml.Linq.XElement SvgConnector(VisioPage page, string id) =>
        System.Xml.Linq.XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { ResolveConnectorLabelOverlaps = false })).Descendants()
            .Single(element => (string?)element.Attribute("data-visio-connector-id") == id);

    // Drawn pixels that turn much lighter once connector labels are drawn: a label backing covers them. Text only darkens.
    private static int LinePixelsCoveredByLabelBackings(VisioPage page) {
        OfficeRasterImage withoutLabels = Raster(page, labels: false), withLabels = Raster(page, labels: true);
        int covered = 0;
        for (int y = 0; y < withLabels.Height; y++) {
            for (int x = 0; x < withLabels.Width; x++) {
                if (Brightness(withLabels.GetPixel(x, y)) - Brightness(withoutLabels.GetPixel(x, y)) > 60) covered++;
            }
        }
        return covered;
    }

    private static OfficeRasterImage Raster(VisioPage page, bool labels) {
        byte[] png = page.ToImage(new VisioImageExportOptions {
            TargetDpi = 96, Supersampling = 1, RenderConnectorLabels = labels, ResolveConnectorLabelOverlaps = false
        }).Bytes;
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? image));
        return image!;
    }

    private static int Brightness(OfficeColor color) => color.R + color.G + color.B;

    [Fact]
    public void WorldMapContinentsWithDegenerateContoursAndInheritedNurbsCellsKeepTheirNativeOutline() {
        // Shape 92 has out-and-back contours that enclose no area; shape 93 has NURBSTo rows whose knots and
        // weights come from the master row and whose E cell is F="Inh" with the NURBS value cached in V.
        VisioDocument document = VisioDocument.Load(Path.Combine(OfficeIMO.Tests.RepositoryTestPaths.Find(), "Assets", "VisioTemplates", "DrawingWithLotsOfShapresAndArrows.vsdx"));
        VisioPage background = document.Pages.Single(page => page.Name == "VBackground-2");

        var projected = background.ToDrawing();

        foreach (string id in new[] { "92", "93" })
            Assert.DoesNotContain(projected.Report.Diagnostics, diagnostic => diagnostic.Code == "VISIO_DRAWING_GEOMETRY_FALLBACK" && diagnostic.Location!.EndsWith(":shape:" + id));
        // Each continent outline is one filled path of many contours, not a four-corner fallback.
        Assert.True(Flatten(projected.Value).Count(shape => shape.FillColor != null && shape.PathCommands.Count > 500) >= 2);
    }

    [Fact]
    public void DiagonalConnectorLinksStayAlongTheRouteInsteadOfItsBoundingBox() {
        VisioDocument document = VisioDocument.Create(); document.UseMastersByDefault = false;
        VisioPage page = document.AddPage("Diagonal", OfficeLength.Inches(8), OfficeLength.Inches(8));
        VisioShape from = page.AddRectangle(OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(0.5), OfficeLength.Inches(0.5), "From");
        VisioShape to = page.AddRectangle(OfficeLength.Inches(7), OfficeLength.Inches(7), OfficeLength.Inches(0.5), OfficeLength.Inches(0.5), "To");
        page.AddConnector("diagonal", from, to, ConnectorKind.Straight).AddHyperlink("https://example.com/route");

        List<OfficeDrawingLink> links = document.ToDrawings().Value[0].Elements.OfType<OfficeDrawingLink>()
            .Where(link => link.Uri == "https://example.com/route").ToList();

        // A single bounding rectangle would cover about 36 square inches of blank page.
        Assert.True(links.Count > 1);
        Assert.All(links, link => Assert.True(link.Width * link.Height < 72D * 72D));
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
