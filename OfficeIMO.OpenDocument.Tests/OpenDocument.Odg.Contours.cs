using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgContourTests {
    [Theory]
    [InlineData("M0 0L160 0L160 60L0 60Z", 0)]
    [InlineData("M0 0L160 0L160 60L0 60Z M20 20", 2)]
    [InlineData("M0 0L160 0L160 60L0 60Z M30 20L130 20M30 40L130 40", 6)]
    [InlineData("M0 0L160 0L160 20L0 20Z L160 60", 4)]
    [InlineData("M0 0M10 10L10 10M0 20L160 20M160 40L160 40", 2)]
    public void NativeShapeClosureAndDrawableContourEndpointsSurviveBothContainers(string data, int markers) {
        var (document, page, path) = Path(data); Bind(path);
        string before = document.GetXml("content.xml").ToString(); string styles = document.GetXml("styles.xml").ToString();
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        foreach (var read in new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) }) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            Assert.Equal(markers, result.Value.Elements.OfType<OfficeDrawingEffectGroup>().Count());
            Assert.Equal(data, read.Pages[0].Shapes[0].PathData);
            Assert.False(result.Report.HasSkippedOrUnsupported);
        }
        Assert.Equal(before, document.GetXml("content.xml").ToString()); Assert.Equal(styles, document.GetXml("styles.xml").ToString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ClosedCompoundFillRetainsHolesAndOpenShapesKeepInactiveFillInXml(bool open) {
        var (document, page, path) = Path("M0 0L160 0L160 60L0 60Z M20 10L20 50L140 50L140 10Z" + (open ? " M30 20L130 20" : ""));
        path.FillColor = OdfColor.Parse("#FFCC88"); Bind(path);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var bitmap = OfficeDrawingRasterRenderer.Render(result.Value);
        Assert.Equal(open ? 0 : 255, bitmap.GetPixel(50, 70).A);
        Assert.Equal(0, bitmap.GetPixel(120, 70).A);
        Assert.Equal(OdfColor.Parse("#FFCC88"), OdgDocument.Load(new MemoryStream(document.ToBytes())).Pages[0].Shapes[0].FillColor);
        Assert.Equal(open, result.Report.ForFeature("shape:Contours:inactive-fill").Any());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DistinctContoursAccumulateOpacityWhileEachMarkersOverlapStaysUniform(bool markers) {
        var (_, page, path) = Path("M0 0L160 60M0 60L160 0");
        path.StrokeWidth = P(6); path.StrokeOpacity = .5; if (markers) Bind(path);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var bitmap = OfficeDrawingRasterRenderer.Render(result.Value);
        Assert.InRange(bitmap.GetPixel(120, 70).A, 189, 194); // Two independent 50% contours overlap.
        Assert.InRange(bitmap.GetPixel(100, 62).A, 125, 130); // One contour only.
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SparseTransparentContoursFitTheRasterBudgetForTheirActualPaint(bool markers) {
        var data = new StringBuilder();
        for (int i = 0; i < 50; i++) data.Append("M0 ").Append(i*5).Append(" L10 ").Append(i*5).Append(' ');
        var (document, page, path) = Path(data.ToString());
        page.Width = page.Height = P(1200);
        path.Bounds = new OdfRect(P(20), P(20), P(1000), P(1000)); path.ViewBox = new OdfViewBox(0, 0, 1000, 1000);
        path.StrokeOpacity = .5; if (markers) Bind(path);
        var bitmap = OfficeDrawingRasterRenderer.Render(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 2_000_000 });
        Assert.InRange(bitmap.GetPixel(25, 20).A, 125, 130);
        Assert.Equal(0, bitmap.GetPixel(800, 800).A);
    }

    [Fact]
    public void MeasurementBudgetCoversAllContoursAndDoesNotClaimPartialMarkerSuccess() {
        var data = new StringBuilder();
        // Each contour fits individually. Together they exceed the bounded operation budget.
        for (int j = 0; j < 2; j++) {
            data.Append("M0 0");
            for (int i = 0; i < 60; i++) data.Append("C100000000 100000000 -100000000 100000000 ").Append(i + 1).Append(" 0");
        }
        var (document, page, path) = Path(data.ToString()); Bind(path);
        string before = document.GetXml("content.xml").ToString(); var result = page.ToDrawing();
        Assert.Equal(2, result.Value.Elements.OfType<OfficeDrawingShape>().Count());
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingEffectGroup>());
        Assert.Equal(OdfConversionMappingStatus.Unsupported, Assert.Single(result.Report.ForFeature("shape:Contours:markers")).Status);
        Assert.Empty(result.Report.ForFeature("shape:Contours:marker-start")); Assert.Empty(result.Report.ForFeature("shape:Contours:marker-end"));
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    private static (OdgDocument Document, OdgPage Page, OdgShape Shape) Path(string data) {
        var document = OdgDocument.Create(); document.Styles.CreateMarker("Arrow", new OdfMarkerGeometry(new OdfViewBox(0, 0, 20, 30), "M10 0L0 30H20Z"));
        var page = document.AddPage("Contours", P(240), P(140));
        var path = page.Shapes.AddPath(new OdfRect(P(40), P(40), P(160), P(60)), new OdfViewBox(0, 0, 160, 60), data, "Contours");
        path.FillColor = null; path.StrokeColor = OdfColor.Parse("#2244BB"); path.StrokeWidth = P(2);
        return (document, page, path);
    }
    private static void Bind(OdgShape path) {
        path.StrokeStartMarkerName = path.StrokeEndMarkerName = "Arrow";
        path.StrokeStartMarkerWidth = path.StrokeEndMarkerWidth = P(20);
    }
    private static OdfLength P(double value) => OdfLength.Points(value);
}
