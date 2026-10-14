using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgMarkerTests {
    private static OdfMarkerGeometry Arrow() => new(new OdfViewBox(0, 0, 20, 30), "M10 0l-10 30h20z");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RoundTripsMarkersAndRendersBothEndpointDirections(bool centered) {
        var document = OdgDocument.Create(); var page = document.AddPage("Markers", P(200), P(100));
        document.Styles.CreateMarker("Arrow", Arrow());
        var line = page.Shapes.AddLine(P(40), P(50), P(160), P(50));
        line.StrokeStartMarkerName = line.StrokeEndMarkerName = "Arrow";
        line.StrokeStartMarkerWidth = line.StrokeEndMarkerWidth = P(20);
        line.StrokeStartMarkerCentered = line.StrokeEndMarkerCentered = centered;
        line.StrokeColor = OdfColor.Parse("#2244BB"); line.StrokeWidth = P(2);
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        foreach (var read in new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) }) {
            Assert.Equal(Arrow().PathData, read.Styles.FindMarker("Arrow")!.Geometry.PathData);
            var shape = read.Pages[0].Shapes[0]; Assert.Equal("Arrow", shape.StrokeStartMarkerName); Assert.Equal(centered, shape.StrokeEndMarkerCentered);
            var scene = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
            var groups = scene.Elements.OfType<OfficeDrawingEffectGroup>().ToArray(); Assert.Equal(2, groups.Length);
            var bitmap = OfficeDrawingRasterRenderer.Render(scene);
            Assert.True(bitmap.GetPixel(centered ? 30 : 45, 50).A > 0); Assert.True(bitmap.GetPixel(centered ? 170 : 155, 50).A > 0);
            Assert.Equal(OfficeColor.Parse("#2244BB"), Assert.Single(groups[0].Drawing.Elements.OfType<OfficeDrawingShape>()).Shape.FillColor);
        }
    }

    [Fact]
    public void InheritedBindingsUseCopyOnWriteAndNullDisablesInheritedMarker() {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", Arrow()); var page = document.AddPage();
        var parent = document.Styles.CreateNamed("MarkerStyle", OdfStyleFamily.Graphic);
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke", "solid");
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-color", "#000000");
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "marker-end", "Arrow");
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "marker-end-width", "20pt");
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "marker-end-center", "1");
        var first = page.Shapes.AddLine(P(10), P(10), P(100), P(10)); first.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", parent.Name);
        var second = page.Shapes.AddLine(P(10), P(40), P(100), P(40)); second.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", parent.Name); document.MarkPartDirty("content.xml");
        first.StrokeEndMarkerCentered = false; first.StrokeEndMarkerCentered = null;
        Assert.True(first.StrokeEndMarkerCentered); first.StrokeEndMarkerName = null;
        Assert.Null(first.StrokeEndMarkerName); Assert.Equal("Arrow", second.StrokeEndMarkerName); Assert.Equal(P(20), second.StrokeEndMarkerWidth);
        Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingEffectGroup>());
    }

    [Fact]
    public void SharedDefinitionEditsRetainMetadataAndScaleActualCurveBounds() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var marker = document.Styles.CreateMarker("Curve", Arrow());
        XElement raw = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "marker").Single(); raw.SetAttributeValue(OdfNamespaces.Draw + "display-name", "Curve label");
        marker.Geometry = new OdfMarkerGeometry(new OdfViewBox(-100, -100, 300, 300), "M0 0Q10 40 20 0Z");
        var line = page.Shapes.AddLine(P(50), P(50), P(150), P(50)); line.StrokeEndMarkerName = "Curve"; line.StrokeEndMarkerWidth = P(20);
        var group = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingEffectGroup>());
        Assert.Equal(20, group.Drawing.Height, 8); // True quadratic maximum is 20, not the control point at 40 or view-box padding.
        Assert.Contains(Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingShape>()).Shape.PathCommands, command => command.Kind == OfficePathCommandKind.QuadraticBezierTo);
        Assert.Equal("Curve label", raw.Attribute(OdfNamespaces.Draw + "display-name")!.Value);
    }

    [Theory]
    [InlineData("M0 0C0 40 20 40 20 0Z", 30)]
    [InlineData("M0 0Q10 40 20 0Z", 20)]
    public void DecimalProducerViewBoxAndCurveExtremaRemainEditable(string path, double height) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        document.Styles.CreateMarker("Decimal", new OdfMarkerGeometry("100 100 20 20.25", path));
        var line = page.Shapes.AddLine(P(30), P(30), P(130), P(30)); line.StrokeEndMarkerName = "Decimal"; line.StrokeEndMarkerWidth = P(20);
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes())); Assert.Equal("100 100 20 21", read.Styles.FindMarker("Decimal")!.Geometry.ViewBox);
        Assert.Equal(height, Assert.Single(read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingEffectGroup>()).Drawing.Height, 8);
    }

    [Fact]
    public void ImportedFractionalViewBoxIsPreservedUntilDefinitionReplacement() {
        var document = OdgDocument.Create(); document.AddPage(); var marker = document.Styles.CreateMarker("Native", Arrow());
        var raw = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "marker").Single();
        raw.SetAttributeValue(OdfNamespaces.Svg + "viewBox", "-0.5 0 20 20.25"); document.MarkPartDirty("styles.xml");
        string before = document.GetXml("styles.xml").ToString(); var geometry = marker.Geometry;
        Assert.Equal("-0.5 0 20 20.25", geometry.ViewBox); Assert.Equal(before, document.GetXml("styles.xml").ToString());
        marker.Geometry = geometry; Assert.Equal("-1 0 21 21", raw.Attribute(OdfNamespaces.Svg + "viewBox")!.Value);
        Assert.Equal(Arrow().PathData, marker.Geometry.PathData);
    }

    [Fact]
    public void TransformedHorizontalLinePaintCanvasIncludesMarkersAndFullStroke() {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", Arrow()); var page = document.AddPage("Translated", P(200), P(100));
        var line = page.Shapes.AddLine(P(40), P(20), P(160), P(20)); line.StrokeColor = OdfColor.Parse("#2244BB"); line.StrokeWidth = P(4);
        line.StrokeStartMarkerName = line.StrokeEndMarkerName = "Arrow"; line.StrokeStartMarkerWidth = line.StrokeEndMarkerWidth = P(20);
        line.Transform = "translate(0pt 30pt)"; var bitmap = OfficeDrawingRasterRenderer.Render(page.ToDrawing().Value);
        Assert.True(bitmap.GetPixel(65, 43).A > 0); Assert.True(bitmap.GetPixel(135, 43).A > 0); Assert.True(bitmap.GetPixel(100, 48).A > 0);
        Assert.Equal(0, bitmap.GetPixel(100, 20).A);
    }

    [Theory]
    [InlineData("M0 0L100 0L100 100Z", 0)]
    [InlineData("M0 0L100 0M0 100L100 100", 2)]
    public void ClosedShapesSuppressMarkersAndOpenContoursEachReceiveTheirBinding(string path, int markers) {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", Arrow()); var page = document.AddPage();
        var shape = page.Shapes.AddPath(new OdfRect(P(30), P(30), P(100), P(100)), new OdfViewBox(0, 0, 100, 100), path);
        shape.StrokeEndMarkerName = "Arrow"; shape.StrokeEndMarkerWidth = P(20);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Equal(markers, result.Value.Elements.OfType<OfficeDrawingEffectGroup>().Count());
        Assert.Equal(markers == 0 ? 1 : 2, result.Value.Elements.OfType<OfficeDrawingShape>().Count());
    }

    [Theory]
    [InlineData("M0 0L100 0L100 100Z")]
    [InlineData("M0 0L100 0M0 100L100 100")]
    public void ZeroWidthMarkerDoesNotRequireEndpointPlacement(string path) {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", Arrow()); var page = document.AddPage();
        var shape = page.Shapes.AddPath(new OdfRect(P(30), P(30), P(100), P(100)), new OdfViewBox(0, 0, 100, 100), path);
        shape.StrokeStartMarkerName = shape.StrokeEndMarkerName = "Arrow"; shape.StrokeStartMarkerWidth = shape.StrokeEndMarkerWidth = P(0);
        Assert.False(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Report.HasSkippedOrUnsupported);
        var rectangle = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)); rectangle.StrokeEndMarkerName = "Arrow"; rectangle.StrokeEndMarkerWidth = P(0);
        Assert.False(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Report.HasSkippedOrUnsupported);
    }

    [Fact]
    public void ConsumedCurveChordAndParentTransformPlaceMarkers() {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", Arrow()); var page = document.AddPage("Curve", P(250), P(250));
        var path = page.Shapes.AddPath(new OdfRect(P(40), P(40), P(100), P(100)), new OdfViewBox(0, 0, 100, 100), "M0 0C0 0 100 0 100 100");
        path.FillColor = null; path.StrokeEndMarkerName = "Arrow"; path.StrokeEndMarkerWidth = P(20); path.Transform = "translate(10pt 20pt)";
        var root = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingEffectGroup>());
        var end = Assert.Single(root.Drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficePoint tip = end.Transform.Then(root.Transform).TransformPoint(new OfficePoint(10, 0));
        Assert.Equal(150, tip.X, 8); Assert.Equal(160, tip.Y, 8);
        OfficePoint baseCenter = end.Transform.Then(root.Transform).TransformPoint(new OfficePoint(10, 30));
        Assert.InRange(baseCenter.X, 145, 149); Assert.InRange(baseCenter.Y, 130, 132);
    }

    [Theory]
    [InlineData("-1pt")]
    [InlineData("200%")]
    [InlineData("+2pt")]
    [InlineData("2PT")]
    public void InvalidEditsAreAtomic(string width) {
        var document = OdgDocument.Create(); var line = document.AddPage().Shapes.AddLine(P(10), P(10), P(100), P(10));
        string before = Snapshot(document);
        Assert.ThrowsAny<ArgumentException>(() => line.StrokeEndMarkerWidth = OdfLength.Parse(width));
        Assert.Throws<ArgumentException>(() => line.StrokeEndMarkerName = "Missing");
        Assert.Throws<ArgumentException>(() => document.Styles.CreateMarker("Not Valid", Arrow()));
        Assert.Throws<ArgumentException>(() => document.Styles.CreateMarker("Bad", new OdfMarkerGeometry(new OdfViewBox(0, 0, 20, 30), "M0 0L0 10")));
        Assert.Equal(before, Snapshot(document));
    }

    [Theory]
    [InlineData("Missing")]
    [InlineData("Duplicate")]
    [InlineData("Path")]
    [InlineData("Width")]
    [InlineData("Center")]
    public void BadImportedMarkerReportsLossAndKeepsUnderlyingLine(string kind) {
        var document = OdgDocument.Create(); var page = document.AddPage(); document.Styles.CreateMarker("Arrow", Arrow());
        var line = page.Shapes.AddLine(P(10), P(10), P(100), P(10)); line.StrokeEndMarkerName = "Arrow"; line.StrokeEndMarkerWidth = P(20);
        var raw = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "marker").Single();
        switch (kind) {
            case "Missing": raw.Remove(); break;
            case "Duplicate": raw.Parent!.Add(new XElement(raw)); break;
            case "Path": raw.SetAttributeValue(OdfNamespaces.Svg + "d", "invalid"); break;
            case "Width": line.StrokeEndMarkerWidth = null; break;
            case "Center": line.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "marker-end-center", "maybe"); break;
        }
        document.MarkPartDirty("styles.xml"); string before = Snapshot(document);
        Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)); Assert.Equal(before, Snapshot(document));
    }

    [Theory]
    [InlineData("start", "")]
    [InlineData("start", " ")]
    [InlineData("end", "")]
    [InlineData("end", " ")]
    public void EmptyImportedWidthRetainsUnderlyingLineAndReportsMarkerLoss(string end, string width) {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", Arrow()); var page = document.AddPage();
        var line = page.Shapes.AddLine(P(10), P(10), P(100), P(10));
        if (end == "start") line.StrokeStartMarkerName = "Arrow"; else line.StrokeEndMarkerName = "Arrow";
        line.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "marker-" + end + "-width", width);
        string before = Snapshot(document); var result = page.ToDrawing();
        Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()); Assert.True(result.Report.HasSkippedOrUnsupported);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)); Assert.Equal(before, Snapshot(document));
    }

    [Fact]
    public void IndependentProducerMarkerRemainsUnchangedWhenReferencedByNewLine() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-network-connectors.odg"));
        var document = OdgDocument.Load(new MemoryStream(OdfTestPackageRewriter.Rewrite(source)));
        var marker = Assert.Single(document.Styles.Markers); Assert.Equal("Arrow", marker.Name); Assert.Equal(Arrow().PathData, marker.Geometry.PathData);
        string before = document.GetXml("styles.xml").ToString(); var line = document.Pages[0].Shapes.AddLine(P(10), P(10), P(100), P(10));
        line.StrokeEndMarkerName = marker.Name; line.StrokeEndMarkerWidth = P(20);
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))); Assert.Equal("Arrow", read.Pages[0].Shapes.Last().StrokeEndMarkerName);
        Assert.Equal(before, read.GetXml("styles.xml").ToString());
    }

    [Fact]
    public void SharedPresentationOwnerAndDisabledStrokePreserveBindings() {
        var document = OdpPresentation.Create(); document.Styles.CreateMarker("Arrow", Arrow());
        var shape = document.AddSlide("Marker").AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)); shape.StrokeEndMarkerName = "Arrow"; shape.StrokeEndMarkerWidth = P(20);
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        Assert.Equal("Arrow", OdpPresentation.LoadFlatXml(flat).Slides[0].Shapes[0].StrokeEndMarkerName);
        var draw = OdgDocument.Create(); draw.Styles.CreateMarker("Arrow", Arrow()); var page = draw.AddPage(); var line = page.Shapes.AddLine(P(10), P(10), P(100), P(10));
        line.StrokeEndMarkerName = "Arrow"; line.StrokeEndMarkerWidth = P(20); line.StrokeColor = null;
        Assert.Empty(page.ToDrawing().Value.Elements.OfType<OfficeDrawingEffectGroup>());
    }
    private static OdfLength P(double n) => OdfLength.Points(n);
    private static string Snapshot(OdfDocument document) => document.GetXml("content.xml").ToString() + document.GetXml("styles.xml");
}
