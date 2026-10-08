using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgMarkerTrimmingTests {
    [Theory]
    [InlineData(10, 1, false)]
    [InlineData(30, 8, false)]
    [InlineData(20, 4, true)]
    public void MarkerWidthAndDockingConsumeStrokeIndependentlyOfStrokeWidth(double width, double strokeWidth, bool center) {
        var (document, page, line) = Line(120);
        line.StrokeWidth = P(strokeWidth); line.StrokeStartMarkerWidth = line.StrokeEndMarkerWidth = P(width);
        line.StrokeStartMarkerCentered = line.StrokeEndMarkerCentered = center;
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var shaft = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()).Shape;
        double consumed = width*1.5*(center ? .5 : 1), trim = consumed - width/15;
        Assert.Equal(trim, shaft.PathCommands[0].Point.X, 8);
        Assert.Equal(120 - trim, shaft.PathCommands[^1].Point.X, 8);
        Assert.Equal(strokeWidth, shaft.StrokeWidth);
        Assert.Equal(2, result.Value.Elements.OfType<OfficeDrawingEffectGroup>().Count());
        Assert.Equal(120, line.X2.ToPoints() - line.X1.ToPoints(), 8); // Projection never edits the source route.
    }

    [Fact]
    public void SemiTransparentMarkerAndShaftShareOneOpacity() {
        var (_, page, line) = Line(120); line.StrokeOpacity = .5;
        var drawing = page.ToDrawing().Value;
        var bitmap = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.InRange(bitmap.GetPixel(68, 50).A, 125, 130); // Deliberate width/15 overlap at the arrow's back.
        Assert.InRange(bitmap.GetPixel(100, 50).A, 125, 130);
        Assert.Equal(.5, Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>()).Opacity);
    }

    [Fact]
    public void CenteredRingHoleDoesNotExposeOriginalShaft() {
        var (document, page, line) = Line(120);
        document.Styles.CreateMarker("Ring", new OdfMarkerGeometry(new OdfViewBox(0, 0, 40, 40),
            "M20 0A20 20 0 1 1 20 40A20 20 0 1 1 20 0Z M20 8A12 12 0 1 0 20 32A12 12 0 1 0 20 8Z"));
        line.StrokeStartMarkerName = "Ring"; line.StrokeStartMarkerCentered = true; line.StrokeEndMarkerName = null;
        var bitmap = OfficeDrawingRasterRenderer.Render(page.ToDrawing().Value);
        Assert.Equal(0, bitmap.GetPixel(40, 50).A); Assert.Equal(0, bitmap.GetPixel(44, 50).A);
        Assert.True(bitmap.GetPixel(51, 50).A > 0);
    }

    [Fact]
    public void NotchedArrowConsumesOnlyItsCenterlineBack() {
        var (document, page, line) = Line(120);
        document.Styles.FindMarker("Arrow")!.Geometry = new OdfMarkerGeometry(new OdfViewBox(0, 0, 20, 30), "M10 0L0 30L10 15L20 30Z");
        var shaft = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>()).Shape;
        Assert.Equal(15 - 20D/15, shaft.PathCommands[0].Point.X, 8);
        Assert.Equal(120 - 15 + 20D/15, shaft.PathCommands[^1].Point.X, 8);
    }

    [Fact]
    public void ExhaustedShortLineKeepsMarkersWithoutResidualStroke() {
        var (_, page, _) = Line(20); var drawing = page.ToDrawing().Value;
        Assert.Empty(drawing.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(2, drawing.Elements.OfType<OfficeDrawingEffectGroup>().Count());
    }

    [Fact]
    public void BadEndMetricDoesNotCancelValidStartOrTrimTheBadEnd() {
        var (document, page, line) = Line(120);
        line.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "marker-end-width", " ");
        string before = document.GetXml("content.xml").ToString(); var result = page.ToDrawing();
        var shaft = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()).Shape;
        Assert.Equal(30 - 20D/15, shaft.PathCommands[0].Point.X, 8); Assert.Equal(120, shaft.PathCommands[^1].Point.X, 8);
        Assert.Single(result.Value.Elements.OfType<OfficeDrawingEffectGroup>()); Assert.True(result.Report.HasSkippedOrUnsupported);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Fact]
    public void OpenCurveRetainsFillXmlWithoutPaintingItAndKeepsExactCurveCommands() {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", new OdfMarkerGeometry(new OdfViewBox(0, 0, 20, 30), "M10 0L0 30H20Z"));
        var page = document.AddPage("Curve", P(200), P(200));
        var path = page.Shapes.AddPath(new OdfRect(P(30), P(30), P(100), P(100)), new OdfViewBox(0, 0, 100, 100), "M0 0Q0 100 100 100");
        path.FillColor = OdfColor.Parse("#88AAEE"); path.StrokeColor = OdfColor.Parse("#2244BB");
        path.StrokeEndMarkerName = "Arrow"; path.StrokeEndMarkerWidth = P(20);
        string before = document.GetXml("content.xml").ToString(); var result = page.ToDrawing();
        var shaft = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()).Shape;
        Assert.Null(shaft.FillColor); Assert.Equal(OfficePathCommandKind.QuadraticBezierTo, shaft.PathCommands[^1].Kind);
        Assert.InRange(shaft.PathCommands[^1].Point.X, 69, 73);
        Assert.Equal(OdfConversionMappingStatus.Converted, Assert.Single(result.Report.ForFeature("shape:" + path.Name + ":inactive-fill")).Status);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
        Assert.Contains("Q", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing().Value));
    }

    [Fact]
    public void MeasurementLimitRetainsSourcePathWithoutClaimingMarkerProjection() {
        var (document, page, _) = Line(120);
        var data = new System.Text.StringBuilder("M0 0");
        for (int i = 0; i < 200; i++) data.Append("C100000000 100000000 -100000000 100000000 ").Append(i + 1).Append(" 0");
        var path = page.Shapes.AddPath(new OdfRect(P(20), P(20), P(100), P(50)), new OdfViewBox(0, 0, 100, 100), data.ToString(), "Bounded");
        path.StrokeEndMarkerName = "Arrow"; path.StrokeEndMarkerWidth = P(20);
        string before = document.GetXml("content.xml").ToString(); var result = page.ToDrawing();
        Assert.Contains(result.Value.Elements.OfType<OfficeDrawingShape>(), shape => shape.Shape.PathCommands.Count == 201);
        Assert.Equal(OdfConversionMappingStatus.Unsupported, Assert.Single(result.Report.ForFeature("shape:Bounded:markers")).Status);
        Assert.Empty(result.Report.ForFeature("shape:Bounded:marker-end"));
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    private static (OdgDocument Document, OdgPage Page, OdgShape Line) Line(double length) {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", new OdfMarkerGeometry(new OdfViewBox(0, 0, 20, 30), "M10 0L0 30H20Z"));
        var page = document.AddPage("Trim", P(200), P(100)); var line = page.Shapes.AddLine(P(40), P(50), P(40 + length), P(50));
        line.StrokeColor = OdfColor.Parse("#2244BB"); line.StrokeWidth = P(2);
        line.StrokeStartMarkerName = line.StrokeEndMarkerName = "Arrow"; line.StrokeStartMarkerWidth = line.StrokeEndMarkerWidth = P(20);
        return (document, page, line);
    }
    private static OdfLength P(double value) => OdfLength.Points(value);
}
