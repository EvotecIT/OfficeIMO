using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgEnhancedGeometryTests {
    [Theory]
    [InlineData(11, 100)]
    [InlineData(21600, 101)]
    public void DeclaredCanvasEdgesRemainExactAfterScalingAndEveryMirror(int canvas, double size) {
        foreach (bool mirrorX in new[] { false, true }) foreach (bool mirrorY in new[] { false, true }) {
            var document = OdgDocument.Create(); var page = document.AddPage();
            var shape = Custom(page, $"M0 0L{canvas} 0 {canvas} {canvas} 0 {canvas}Z N", new OdfViewBox(0, 0, canvas, canvas));
            shape.Bounds = new OdfRect(P(10), P(20), P(size), P(size));
            var native = shape.Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!;
            native.SetAttributeValue(OdfNamespaces.Draw + "mirror-horizontal", mirrorX);
            native.SetAttributeValue(OdfNamespaces.Draw + "mirror-vertical", mirrorY); document.MarkPartDirty("content.xml");
            foreach (var read in RoundTrips(document)) {
                var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
                var projected = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()).Shape;
                Assert.Equal(new OfficePoint(mirrorX ? size : 0, mirrorY ? size : 0), projected.PathCommands[0].Point);
                Assert.Equal(new OfficePoint(mirrorX ? 0 : size, mirrorY ? 0 : size), projected.PathCommands[2].Point);
                Assert.All(projected.PathCommands.Where(command => command.Kind != OfficePathCommandKind.Close), command => {
                    Assert.InRange(command.Point.X, 0, size); Assert.InRange(command.Point.Y, 0, size);
                });
            }
        }
    }

    [Fact]
    public void ExplicitAndRepeatedCubicCommandsBothReachTheNormalizedCommandLimit() {
        OfficePathCommand[]? expected = null;
        foreach (string path in new[] {
            "M0 0" + string.Concat(Enumerable.Repeat("C0 0 1 1 2 2", 19999)) + " N",
            "M0 0 C" + string.Concat(Enumerable.Repeat("0 0 1 1 2 2 ", 19999)) + " N"
        }) {
            var document = OdgDocument.Create(); var page = document.AddPage();
            Custom(page, path, new OdfViewBox(0, 0, 100, 100));
            var projected = Assert.Single(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingShape>());
            var commands = projected.Shape.PathCommands.ToArray(); Assert.Equal(20000, commands.Length);
            if (expected != null) Assert.Equal(expected, commands);
            expected = commands;
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ImportedLiteralGeometryScalesAndMirrorsWithoutReplacingItsNativePath(bool mirrorX, bool mirrorY) {
        var document = OdgDocument.Create(); var page = document.AddPage("Enhanced", P(240), P(180));
        var shape = Custom(page, "M-100 50L300 50 300 250Z N", new OdfViewBox(-100, 50, 400, 200));
        shape.Bounds = new OdfRect(P(10), P(20), P(160), P(40));
        var native = shape.Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!;
        native.SetAttributeValue(OdfNamespaces.Draw + "mirror-horizontal", mirrorX ? "1" : "false");
        native.SetAttributeValue(OdfNamespaces.Draw + "mirror-vertical", mirrorY ? "true" : "0");
        string expected = native.ToString(); document.MarkPartDirty("content.xml");
        foreach (var read in RoundTrips(document)) {
            string before = read.GetXml("content.xml").ToString();
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var projected = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
            Assert.Equal(10, projected.X); Assert.Equal(20, projected.Y);
            Assert.Equal(new OfficePoint(mirrorX ? 160 : 0, mirrorY ? 40 : 0), projected.Shape.PathCommands[0].Point);
            Assert.Equal(new OfficePoint(mirrorX ? 0 : 160, mirrorY ? 0 : 40), projected.Shape.PathCommands[2].Point);
            Assert.Equal(expected, read.Pages[0].Shapes[0].Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!.ToString());
            Assert.Equal(before, read.GetXml("content.xml").ToString());
        }
    }

    [Fact]
    public void OneEnhancedPaintSetUsesEvenOddFillForAllItsSubpaths() {
        var document = OdgDocument.Create(); var page = document.AddPage("Hole", P(240), P(180));
        var shape = Custom(page, "M0 0L100 0 100 100 0 100Z M25 25L75 25 75 75 25 75Z N", new OdfViewBox(0, 0, 100, 100));
        shape.Bounds = new OdfRect(P(10), P(10), P(100), P(100)); shape.FillRule = OfficeFillRule.NonZero;
        shape.FillColor = OdfColor.Parse("00ff00"); shape.StrokeColor = null;
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var geometry = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()).Shape;
            Assert.Equal(OfficeFillRule.EvenOdd, geometry.FillRule);
            var pixels = OfficeDrawingRasterRenderer.Render(result.Value);
            Assert.Equal(OfficeColor.FromRgb(0, 255, 0), pixels.GetPixel(20, 20));
            Assert.Equal(OfficeColor.Transparent, pixels.GetPixel(60, 60));
        }
    }

    [Fact]
    public void EnhancedLiteralCurvesUseTheSharedPathGrammarAndDeclaredCanvas() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = Custom(page, "M0,0 Q25 50 50 0 C50 -1e1 75 -10 100 0 N", new OdfViewBox(0, -10, 100, 60));
        shape.Bounds = new OdfRect(P(10), P(20), P(100), P(60));
        var path = Assert.Single(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingShape>()).Shape.PathCommands;
        Assert.Equal(OfficePathCommandKind.QuadraticBezierTo, path[1].Kind);
        Assert.Equal(new OfficePoint(25, 60), path[1].ControlPoint1);
        Assert.Equal(OfficePathCommandKind.CubicBezierTo, path[2].Kind);
        Assert.Equal(new OfficePoint(50, 0), path[2].ControlPoint1);
    }

    [Theory]
    [InlineData(0, 360)]
    [InlineData(45, 405)]
    [InlineData(45, -315)]
    public void FullEnhancedEllipsesReuseSharedArcGeometry(double start, double end) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = Custom(page, FormattableString.Invariant($"U100 50 100 50 {start} {end} Z N"), new OdfViewBox(0, 0, 200, 100));
        shape.Bounds = new OdfRect(P(10), P(20), P(200), P(100));
        foreach (var read in RoundTrips(document)) {
            var geometry = Assert.Single(read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingShape>()).Shape;
            Assert.Contains(geometry.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);
            Assert.All(geometry.PathCommands.Where(command => command.Kind == OfficePathCommandKind.CubicBezierTo), command => {
                foreach (OfficePoint point in new[] { command.Point, command.ControlPoint1, command.ControlPoint2 }) {
                    Assert.InRange(point.X, 0, geometry.Width); Assert.InRange(point.Y, 0, geometry.Height);
                }
            });
            Assert.Equal(OfficePathCommandKind.Close, geometry.PathCommands.Last().Kind);
            var points = geometry.PathCommands.Where(command => command.Kind != OfficePathCommandKind.Close).Select(command => command.Point).ToArray();
            Assert.All(points, point => Assert.Equal(1, Math.Pow((point.X - 100) / 100, 2) + Math.Pow((point.Y - 50) / 50, 2), 10));
            Assert.Equal(points[0].X, points.Last().X, 10); Assert.Equal(points[0].Y, points.Last().Y, 10);
        }
    }

    [Theory]
    [InlineData("M0 0L?f0 100Z N")]
    [InlineData("M0 0L$0 100Z N")]
    [InlineData("M0 0L100 100 F N")]
    [InlineData("M0 0L100 100 S N")]
    [InlineData("M0 0L100 100 N M0 100L100 0 N")]
    [InlineData("U50 50 50 50 0 180 Z N")]
    [InlineData("M0 0L100 100")]
    [InlineData("M0 0L1e309 100 N")]
    [InlineData("M0 0Q1 2 N")]
    [InlineData("m0 0l100 100 N")]
    [InlineData("M0 0L101 100 N")]
    [InlineData("M0 0L100.00000000000001 100 N")]
    [InlineData("M0 0C0 0 100.00000000000001 100 100 100 N")]
    public void UnsupportedEnhancedGeometryKeepsTextAndReportsLossWithoutSubstitutingAPreset(string data) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = Custom(page, data, new OdfViewBox(0, 0, 100, 100));
        shape.Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!.SetAttributeValue(OdfNamespaces.Draw + "type", "rectangle");
        shape.Text = "Retained caption";
        foreach (var read in RoundTrips(document)) {
            string before = read.GetXml("content.xml").ToString(); var result = read.Pages[0].ToDrawing();
            Assert.Empty(result.Value.Elements.OfType<OfficeDrawingShape>());
            Assert.Equal("Retained caption", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Message!.Contains("custom-shape"));
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.Equal(data, (string?)read.Pages[0].Shapes[0].Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!.Attribute(OdfNamespaces.Draw + "enhanced-path"));
            Assert.Equal(before, read.GetXml("content.xml").ToString());
        }
    }

    [Theory]
    [InlineData("path-stretchpoint-x", "50")]
    [InlineData("path-stretchpoint-y", "50")]
    [InlineData("extrusion", "true")]
    [InlineData("mirror-horizontal", "invalid")]
    public void UnsupportedGeometryEffectsDoNotProduceAnOrdinaryRectangle(string name, string value) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = Custom(page, "M0 0L100 100 N", new OdfViewBox(0, 0, 100, 100));
        shape.Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!.SetAttributeValue(OdfNamespaces.Draw + name, value);
        Assert.Empty(page.ToDrawing().Value.Elements);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void EnhancedTextAreaLimitIsReportedAlongsideSupportedGeometryAndCaption() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = Custom(page, "M0 0L100 0 100 100Z N", new OdfViewBox(0, 0, 100, 100)); shape.Text = "Caption";
        shape.Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!.SetAttributeValue(OdfNamespaces.Draw + "text-areas", "25 25 75 75");
        var result = page.ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()); Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature.EndsWith(":enhanced-text-area", StringComparison.Ordinal) && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void NativeEmptyParagraphsDoNotTurnUnusedEnhancedTextAreasIntoAConversionLoss() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = Custom(page, "U50 50 50 50 0 360 Z N", new OdfViewBox(0, 0, 100, 100));
        shape.Element.Add(new XElement(OdfNamespaces.Text + "p"));
        shape.Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!.SetAttributeValue(OdfNamespaces.Draw + "text-areas", "25 25 75 75");
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
            Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature.EndsWith(":enhanced-text-area", StringComparison.Ordinal));
        }
    }

    [Theory]
    [InlineData("text-areas", "25 25 75 75", false)]
    [InlineData("text-rotate-angle", "45", false)]
    [InlineData("text-path", "true", false)]
    [InlineData("text-areas", "25 25 75 75", true)]
    public void LabelOnlyCaptionsStillReportEnhancedTextLayoutLoss(string attribute, string value, bool numbered) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = Custom(page, "M0 0L100 0 100 100Z N", new OdfViewBox(0, 0, 100, 100));
        shape.AddList(numbered).AddItem("");
        shape.Element.Element(OdfNamespaces.Draw + "enhanced-geometry")!.SetAttributeValue(OdfNamespaces.Draw + attribute, value);
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
            var caption = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
            Assert.False(string.IsNullOrEmpty(Assert.Single(caption.Paragraphs).Label?.Run.Text));
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature.EndsWith(":enhanced-text-area", StringComparison.Ordinal) && mapping.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        }
    }

    [Fact]
    public void IndependentMasterAndPageLiteralCustomShapesProjectWhileFormulaArtworkRemainsReported() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-master-artwork.odg"));
        foreach (var read in RoundTrips(document)) {
            string[] before = new[] { read.GetXml("content.xml").ToString(), read.GetXml("styles.xml").ToString() };
            var result = read.Pages[0].ToDrawing();
            Assert.Equal(6, result.Value.Elements.OfType<OfficeDrawingShape>().Count());
            Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Message!.Contains("modifier references"));
            Assert.Equal(before, new[] { read.GetXml("content.xml").ToString(), read.GetXml("styles.xml").ToString() });
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EnhancedInputLimitsRejectExcessiveDataWithoutChangingTheSource(bool textLimit) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        string path = textLimit ? new string(' ', 1024 * 1024 + 1) : "M0 0" + string.Concat(Enumerable.Repeat("L1 1", 20000)) + " N";
        var shape = Custom(page, path, new OdfViewBox(0, 0, 100, 100));
        string before = shape.ToXml().ToString(); var result = page.ToDrawing();
        Assert.Empty(result.Value.Elements);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Message!.Contains("exceeds"));
        Assert.Equal(before, shape.ToXml().ToString());
    }

    private static OdfLength P(double value) => OdfLength.Points(value);
    private static OdgShape Custom(OdgPage page, string path, OdfViewBox box) {
        var shape = page.Shapes.AddRectangle(new OdfRect(P(10), P(20), P(80), P(40)), "Enhanced");
        shape.Element.Name = OdfNamespaces.Draw + "custom-shape";
        shape.Element.Add(new XElement(OdfNamespaces.Draw + "enhanced-geometry", new XAttribute(OdfNamespaces.Svg + "viewBox", box),
            new XAttribute(OdfNamespaces.Draw + "enhanced-path", path)));
        shape.Document.MarkPartDirty("content.xml"); return shape;
    }
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        byte[] package = document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource });
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(package)), OdgDocument.LoadFlatXml(flat) };
    }
}
