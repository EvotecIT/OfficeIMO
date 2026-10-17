using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartSeriesPaintTests {
    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.Bubble)]
    public void PointPaintOverridesInheritFillAndOutlineIndependently(OfficeChartKind kind) {
        var series = kind == OfficeChartKind.Bubble
            ? OfficeChartSeries.CreateBubble("Paint", new[] { 1d, 2d, 3d }, new[] { 4d, 5d, 6d }, new[] { 1d, 2d, 3d })
            : new OfficeChartSeries("Paint", new[] { 4d, 5d, 6d });
        var paint = new OfficeChartPointStyle(noFill: true, outlineColor: OfficeColor.Red, outlineWidth: 3, showOutline: true);
        series = series.WithSeriesStyle(paint).WithPointStyles(new OfficeChartPointStyle?[] {
            new(outlineColor: OfficeColor.Blue), new(fillColor: OfficeColor.Green), new(showOutline: false)
        }).WithLegendVisibility(false);
        Assert.Same(paint, series.SeriesStyle);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Paint", null, kind,
            new OfficeChartData(new[] { "A", "B", "C" }, new[] { series }), 640, 360,
            layout: new OfficeChartLayout(showLegend: false)));
        var points = drawing.Shapes.Where(shape => shape.Shape.Kind == (kind == OfficeChartKind.Bubble
            ? OfficeShapeKind.Ellipse : OfficeShapeKind.Rectangle) && shape.Shape.StrokeWidth == 3).ToArray();
        Assert.Contains(points, point => point.Shape.FillColor == null && point.Shape.StrokeColor == OfficeColor.Blue);
        Assert.Contains(points, point => point.Shape.FillColor == OfficeColor.Green && point.Shape.StrokeColor == OfficeColor.Red);
        Assert.DoesNotContain(drawing.Shapes, point => point.Shape.FillColor == OfficeColor.Red);
    }
}
