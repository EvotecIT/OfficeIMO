using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Fact]
    public void FreeLabelledConnectorsDeclareTheRequiredOdfCoordinateCanvas() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(40, 40), new OfficePoint(160, 80));
        AddLabel(document, connector);
        foreach (var read in RoundTrips(document))
            Assert.Equal("0 0 1 1", (string?)read.Pages[0].Shapes[0].ToXml().Attribute(OdfNamespaces.Svg + "viewBox"));
    }

    [Fact]
    public void ProjectsIndependentPartiallyAttachedLabelWithoutChangingProducerXml() {
        var document = LoadProducer();
        var connector = document.Pages[0].Shapes.Single(shape => shape.XmlId == "id13");
        string source = document.GetXml("content.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var frames = Frames(read.Pages[0].ToDrawing().Value);
            var label = Assert.Single(frames, frame => frame.Text.PlainText == "95.217.96.40/29\n95.216.28.178");
            // The unchanged producer SVG centers both lines at native x=2609 (1/100 mm).
            Assert.Equal(2609 * 72D / 2540, Center(label).X, 6);
            Assert.Equal((4241 + 7200) / 2D * 72 / 2540, Center(label).Y, 6);
            Assert.All(label.Text.Paragraphs, p => Assert.Equal(OfficeTextAlignment.Center, p.Alignment));
            Assert.All(label.Text.Paragraphs.SelectMany(p => p.Runs), run => Assert.Equal(10, run.FontSize));
            Assert.False(label.Text.WrapText);
            Assert.Equal(OfficeTextVerticalAlignment.Center, label.Text.VerticalAlignment);
        }
        Assert.Equal(source, document.GetXml("content.xml").ToString());
        Assert.Equal("id1", connector.StartShapeId); Assert.Null(connector.EndShapeId);
    }

    [Theory]
    [InlineData(false, 50, 50, 150, 50)]
    [InlineData(true, 50, 50, 150, 50)]
    [InlineData(false, 100, 30, 100, 160)]
    [InlineData(true, 100, 30, 100, 160)]
    [InlineData(false, 40, 40, 160, 120)]
    [InlineData(false, 160, 120, 40, 40)]
    [InlineData(false, 150, 50, 50, 50)]
    [InlineData(true, 160, 120, 40, 40)]
    public void ProjectsFullStyledUnwrappedLabelsOnHorizontalVerticalAndDiagonalLines(bool connector,
        double x1, double y1, double x2, double y2) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var line = connector ? page.Shapes.AddConnector(new OfficePoint(x1, y1), new OfficePoint(x2, y2)) :
            page.Shapes.AddLine(OdfLength.Points(x1), OdfLength.Points(y1), OdfLength.Points(x2), OdfLength.Points(y2));
        AddLabel(document, line);
        string source = document.GetXml("content.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var label = Assert.Single(Frames(result.Value));
            Assert.Equal("Primary emphasized\nSecond line", label.Text.PlainText);
            Assert.Equal((x1 + x2) / 2, Center(label).X, 6);
            Assert.Equal((y1 + y2) / 2, Center(label).Y, 6);
            double angle = connector ? 0 : Math.Atan2(y2 - y1, x2 - x1);
            Assert.Equal(Math.Cos(angle), label.Transform.M11, 6);
            Assert.Equal(Math.Sin(angle), label.Transform.M12, 6);
            var emphasized = label.Text.Paragraphs[0].Runs.Single(run => run.Text == "emphasized");
            Assert.True(emphasized.Bold); Assert.Equal(OfficeColor.Parse("#BA1234"), emphasized.Color);
            string svg = OfficeDrawingSvgExporter.ToSvg(result.Value);
            Assert.Contains("emphasized", svg); Assert.Contains("Second line", svg);
            Assert.DoesNotContain("…", svg);
        }
        Assert.Equal(source, document.GetXml("content.xml").ToString());
    }

    [Fact]
    public void LabelsUseSavedRouteBoundsRatherThanTheEndpointMidpoint() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(40, 40), new OfficePoint(160, 80), OdgConnectorKind.Lines);
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(40, 40), OfficePathCommand.LineTo(40, 180),
            OfficePathCommand.LineTo(160, 180), OfficePathCommand.LineTo(160, 80) });
        AddLabel(document, connector);
        var frame = Assert.Single(Frames(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value));
        Assert.InRange(Center(frame).X, 99.98, 100.02);
        Assert.InRange(Center(frame).Y, 109.98, 110.02);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TranslationBakingRetainsTextStylesAndLabelPlacementAcrossReopening(bool groupOperation) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var group = page.Shapes.AddGroup();
        var connector = group.Children.AddConnector(new OfficePoint(40, 40), new OfficePoint(160, 80));
        AddLabel(document, connector);
        string styles = document.GetXml("styles.xml").ToString();
        if (groupOperation) group.TransformChildren("translate(20pt 30pt)");
        else { connector.Transform = "translate(20pt 30pt)"; connector.BakeConnectorTransform(); }
        Assert.Equal(styles, document.GetXml("styles.xml").ToString());
        foreach (var read in RoundTrips(document)) {
            var frame = Assert.Single(Frames(read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value));
            Assert.InRange(Center(frame).X, 119.98, 120.02);
            Assert.InRange(Center(frame).Y, 89.98, 90.02);
            Assert.All(frame.Text.Paragraphs.SelectMany(p => p.Runs), run => Assert.Equal(12, run.FontSize));
            Assert.Null(read.Pages[0].Shapes[0].Children[0].Transform);
        }
    }

    [Fact]
    public void LabelsAtThePageEdgeKeepTheirAnchorAndTextInsteadOfBeingDropped() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(-10, 0), new OfficePoint(10, 0));
        AddLabel(document, connector);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var frame = Assert.Single(Frames(result.Value));
        Assert.Equal(0, Center(frame).X, 6); Assert.Equal(0, Center(frame).Y, 6);
        Assert.Contains("Second line", OfficeDrawingSvgExporter.ToSvg(result.Value));
    }

    [Fact]
    public void RotatedCondensedLabelsKeepReportOnlyInkAndAnchorsAndRejectStrictConversion() {
        var horizontal = Caption(false); var vertical = Caption(true);
        var horizontalFrame = Assert.Single(Frames(horizontal)); var verticalFrame = Assert.Single(Frames(vertical));
        Assert.Equal(Center(horizontalFrame).X, Center(verticalFrame).X, 6);
        Assert.Equal(Center(horizontalFrame).Y, Center(verticalFrame).Y, 6);
        Assert.Equal(8, horizontalFrame.Text.Height, 6); Assert.Equal(8, verticalFrame.Text.Height, 6);
        int expected = Ink(horizontal), actual = Ink(vertical);
        Assert.True(expected > 50);
        Assert.InRange(actual, expected - 2, expected + 2);

        static OfficeDrawing Caption(bool rotated) {
            var document = OdgDocument.Create(); var page = document.AddPage();
            page.Width = page.Height = OdfLength.Points(200);
            var line = page.Shapes.AddLine(OdfLength.Points(rotated ? 100 : 50), OdfLength.Points(rotated ? 50 : 100),
                OdfLength.Points(rotated ? 100 : 150), OdfLength.Points(rotated ? 150 : 100));
            line.StrokeColor = null;
            var paragraph = line.AddParagraph("Condensed glyphs");
            paragraph.FontSize = OdfLength.Points(12); paragraph.LineHeight = OdfLength.Points(8);
            paragraph.TextAlign = "center"; paragraph.Color = OdfColor.Parse("#FF0000");
            Graphic(document, line).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
            var result = page.ToDrawing();
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal) &&
                m.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            return result.Value;
        }

        static int Ink(OfficeDrawing drawing) {
            var image = OfficeDrawingRasterRenderer.Render(drawing); int count = 0;
            for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.A > 0 && pixel.R > 200 && pixel.G < 50 && pixel.B < 50) count++;
            }
            return count;
        }
    }

    [Theory]
    [InlineData(false, false, false, false)]
    [InlineData(false, true, false, false)]
    [InlineData(false, false, true, false)]
    [InlineData(true, true, false, false)]
    [InlineData(false, false, false, true)]
    [InlineData(false, false, true, true)]
    public void GraphicWritingDirectionsAreReportedForLinesAndOrdinaryShapeText(bool rectangle, bool inherited, bool pageDirection, bool paragraphDefault) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = rectangle ? page.Shapes.AddRectangle(new OdfRect(OdfLength.Points(40), OdfLength.Points(40),
            OdfLength.Points(100), OdfLength.Points(40))) : page.Shapes.AddConnector(new OfficePoint(40, 40), new OfficePoint(160, 80));
        AddLabel(document, shape);
        var style = Graphic(document, shape);
        if (inherited) {
            var parent = document.Styles.CreateNamed("VerticalCaption", OdfStyleFamily.Graphic);
            style.ParentStyleName = parent.Name; style = parent;
        }
        style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Style + "writing-mode", pageDirection ? "page" : "tb-rl");
        if (pageDirection) document.GetXml("styles.xml").Descendants(OdfNamespaces.Style + "page-layout-properties").Single()
            .SetAttributeValue(OdfNamespaces.Style + "writing-mode", "tb-rl");
        if (paragraphDefault) document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
            new System.Xml.Linq.XElement(OdfNamespaces.Style + "default-style",
                new System.Xml.Linq.XAttribute(OdfNamespaces.Style + "family", "paragraph"),
                new System.Xml.Linq.XElement(OdfNamespaces.Style + "paragraph-properties",
                    new System.Xml.Linq.XAttribute(OdfNamespaces.Style + "writing-mode", "lr-tb"))));
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing();
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":writing-mode", StringComparison.Ordinal) &&
                m.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        }
    }

    [Fact]
    public void ParagraphPageDirectionInheritsItsHorizontalGraphicBeforeThePageLayout() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(40, 40), new OfficePoint(160, 80));
        AddLabel(document, connector);
        foreach (var paragraph in connector.Paragraphs) paragraph.WritingMode = "page";
        Graphic(document, connector).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Style + "writing-mode", "lr-tb");
        document.GetXml("styles.xml").Descendants(OdfNamespaces.Style + "page-layout-properties").Single()
            .SetAttributeValue(OdfNamespaces.Style + "writing-mode", "tb-rl");
        foreach (var read in RoundTrips(document)) Assert.Single(Frames(read.Pages[0]
            .ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value));
    }

    [Theory]
    [InlineData("rotated-curve")]
    [InlineData("invalid-text-area-alignment")]
    [InlineData("invalid-wrap")]
    public void UnsupportedLabelProfilesKeepGeometryAndAnActionableLoss(string profile) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(40, 40), new OfficePoint(160, 80));
        connector.Name = "Label"; AddLabel(document, connector);
        switch (profile) {
            case "rotated-curve":
                connector.ConnectorKind = OdgConnectorKind.Curve;
                connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(40, 40),
                    OfficePathCommand.CubicBezierTo(40, 100, 160, 100, 160, 80) });
                connector.Transform = "rotate(0.2) translate(80pt 40pt)"; break;
            case "invalid-text-area-alignment": Graphic(document, connector).SetProperty(OdfNamespaces.Style + "graphic-properties",
                OdfNamespaces.Draw + "textarea-horizontal-align", "invalid"); break;
            case "invalid-wrap": Graphic(document, connector).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "invalid"); break;
        }
        string before = document.GetXml("content.xml").ToString();
        var result = page.ToDrawing(); Assert.Empty(Frames(result.Value));
        Assert.Contains(result.Report.Mappings, m => m.Feature == "shape:Label:text" && m.Status == OdfConversionMappingStatus.Skipped);
        Assert.Contains("stroke", OfficeDrawingSvgExporter.ToSvg(result.Value));
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    private static void AddLabel(OdgDocument document, OdgShape shape) {
        shape.FontFamily = "Liberation Sans";
        var p = shape.AddParagraph("Primary "); p.FontSize = OdfLength.Points(12); p.TextAlign = "center";
        var run = p.AddRun("emphasized"); run.Bold = true; run.Color = OdfColor.Parse("#BA1234");
        var second = shape.AddParagraph("Second line"); second.FontSize = OdfLength.Points(12); second.TextAlign = "center";
        Graphic(document, shape).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        Graphic(document, shape).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-vertical-align", "middle");
    }

    private static OdfStyle Graphic(OdgDocument document, OdgShape shape) =>
        document.Styles.Find(OdfStyleFamily.Graphic, (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!)!;

    private static OdgDocument LoadProducer() => OdgDocument.Load(new MemoryStream(OdfTestPackageRewriter.Rewrite(
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-network-connectors.odg")))));

    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        yield return document;
        yield return OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource })));
        using var flat = new MemoryStream();
        document.SaveFlatXml(flat);
        flat.Position = 0;
        yield return OdgDocument.LoadFlatXml(flat);
    }

    private static OfficePoint Center((OfficeDrawingRichText Text, double X, double Y, OfficeTransform Transform) frame) =>
        frame.Transform.TransformPoint(new OfficePoint(frame.Text.X + frame.Text.Width / 2, frame.Text.Y + frame.Text.Height / 2));

    private static List<(OfficeDrawingRichText Text, double X, double Y, OfficeTransform Transform)> Frames(OfficeDrawing drawing) {
        var result = new List<(OfficeDrawingRichText, double, double, OfficeTransform)>();
        Visit(drawing, OfficeTransform.Identity); return result;
        void Visit(OfficeDrawing current, OfficeTransform transform) {
            foreach (var element in current.Elements) {
                if (element is OfficeDrawingRichText text) {
                    var point = transform.TransformPoint(new OfficePoint(text.X, text.Y)); result.Add((text, point.X, point.Y, transform));
                } else if (element is OfficeDrawingEffectGroup group) Visit(group.Drawing, group.Transform.Then(transform));
            }
        }
    }
}
