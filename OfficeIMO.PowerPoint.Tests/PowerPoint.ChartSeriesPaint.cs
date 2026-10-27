using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartSeriesPaintTests {
    [Fact]
    public void NativeSeriesFillAndOutlineRemainIndependent() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "Charts", "SeriesPaint", "series-paint.pptx");
        for (int number = 0; number < 16; number++) {
            using var document = PowerPointPresentation.Load(path);
            var chart = document.Slides[number].Charts.Single();
            Assert.True(chart.TryGetSnapshot(out var native, out var reason), reason);
            var snapshot = PowerPointChartSnapshotMapper.ToOfficeSnapshot(native, 640, 360);
            CheckPaint(snapshot!, number);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeRoundTripKeepsMarkerFillIndependentOfConnectingLine(bool pointOverride) {
        var series = new OfficeChartSeries("Paint", new[] { 2d, 4d }, null, OfficeColor.Blue,
            null, true, markerSize: 10, markerShape: OfficeChartMarkerShape.Circle,
            strokeDashStyle: OfficeStrokeDashStyle.Dash)
            .WithSeriesStyle(new OfficeChartPointStyle(outlineColor: OfficeColor.Red, outlineWidth: 3));
        if (pointOverride) series = series.WithPointStyles(new OfficeChartPointStyle?[] { new(fillColor: OfficeColor.Green), null });
        using var document = PowerPointPresentation.Create();
        document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" }, new[] { series }));
        using var stream = new MemoryStream(document.ToBytes());
        using var reopened = PowerPointPresentation.Load(stream);
        Assert.Empty(reopened.ValidateDocument());
        Assert.True(reopened.Slides[0].Charts.Single().TryGetOfficeSnapshot(out var snapshot));
        var drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Contains(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Line &&
            shape.Shape.StrokeColor == OfficeColor.Red && shape.Shape.StrokeWidth == 3 && shape.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);
        Assert.Contains(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Ellipse && shape.Shape.FillColor == OfficeColor.Blue);
        Assert.Equal(pointOverride, drawing.Shapes.Any(shape => shape.Shape.Kind == OfficeShapeKind.Ellipse && shape.Shape.FillColor == OfficeColor.Green));
        Assert.DoesNotContain(reopened.ConvertToPdfWithEvidence().Warnings, warning => warning.Code == "unsupported-chart");
    }

    private static void CheckPaint(OfficeChartSnapshot snapshot, int number) {
        bool fill = number % 4 < 2, outline = number % 2 == 0, line = number >= 12;
        var paint = Assert.Single(snapshot.Data.Series).SeriesStyle;
        Assert.NotNull(paint);
        Assert.Equal(!fill, paint.NoFill);
        Assert.Equal(outline, paint.ShowOutline);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Equal(fill && !line, drawing.Shapes.Any(shape => shape.Shape.FillColor == OfficeColor.Parse("#156082")));
        Assert.Equal(outline, drawing.Shapes.Any(shape => shape.Shape.StrokeColor == OfficeColor.Parse("#E97132") && shape.Shape.StrokeWidth == 3));
    }

    [Theory]
    [InlineData(OfficeChartKind.Area, false)]
    [InlineData(OfficeChartKind.Area, true)]
    [InlineData(OfficeChartKind.Bubble, false)]
    [InlineData(OfficeChartKind.Bubble, true)]
    [InlineData(OfficeChartKind.ColumnClustered, false)]
    [InlineData(OfficeChartKind.ColumnClustered, true)]
    [InlineData(OfficeChartKind.Line, false)]
    [InlineData(OfficeChartKind.Line, true)]
    public void NativeColourSetterKeepsExplicitOutlineAfterRemovingAutomaticFill(OfficeChartKind kind, bool byName) {
        var series = kind == OfficeChartKind.Bubble
            ? OfficeChartSeries.CreateBubble("Paint", new[] { 1d, 2d }, new[] { 2d, 4d }, new[] { 3d, 5d })
            : new OfficeChartSeries("Paint", new[] { 2d, 4d }, null, null, null, showMarkers: false);
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] { series }));
        if (byName) chart.SetSeriesLineColor("Paint", "FF0000", 3);
        else chart.SetSeriesLineColor(0, "FF0000", 3);
        var properties = document.Slides[0].SlidePart.ChartParts.Single().ChartSpace!
            .Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>().First();
        properties.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.SolidFill>();
        properties.AddChild(new DocumentFormat.OpenXml.Drawing.NoFill(), true);
        Assert.True(chart.TryGetSnapshot(out var native, out var reason), reason);
        var snapshot = PowerPointChartSnapshotMapper.ToOfficeSnapshot(native, 640, 360,
            new OfficeChartStyle(palette: new[] { OfficeColor.Green }));
        Assert.True(snapshot.Data.Series[0].SeriesStyle!.NoFill);
        Assert.Equal(OfficeColor.Red, snapshot.Data.Series[0].SeriesStyle!.OutlineColor);
        Assert.Contains(OfficeChartDrawingRenderer.Render(snapshot).Shapes,
            shape => shape.Shape.StrokeColor == OfficeColor.Red && shape.Shape.StrokeWidth == 3);
    }
}
