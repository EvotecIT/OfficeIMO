using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartRadialLayoutTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    [InlineData(OfficeChartKind.Radar)]
    public void SmallAuthoredCanvas_KeepsRadialGeometryInsideTheFrame(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B", "C" },
            new[] { new OfficeChartSeries("Values", new[] { 3d, 4d, 5d }) });
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            kind, data, 90, 30, null, new OfficeChartLayout(showLegend: false)), false);
        var polygons = drawing.Shapes.Where(shape => shape.Shape.Kind == OfficeShapeKind.Polygon).ToArray();
        Assert.NotEmpty(polygons);
        foreach (var polygon in polygons)
            foreach (var point in polygon.Shape.Points) {
                Assert.InRange(polygon.X + point.X, -0.000001, 90.000001);
                Assert.InRange(polygon.Y + point.Y, -0.000001, 30.000001);
            }
    }

    [Theory]
    [InlineData(10)]
    [InlineData(50)]
    [InlineData(90)]
    public void RadialGeometry_HoleMatchesPercentageAndLabelAnchorStaysOnRing(int hole) {
        var data = new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Status", new[] { 7d }) });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Status", null,
            OfficeChartKind.Doughnut, data, 400, 400, null,
            new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true),
            new OfficeChartRadialLayout(90, hole)));
        OfficeDrawingShape ring = Assert.Single(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Polygon);
        double centerX = ring.X + ring.Shape.Width / 2;
        double centerY = ring.Y + ring.Shape.Height / 2;
        double Radius(OfficePoint point) => Math.Sqrt(Math.Pow(ring.X + point.X - centerX, 2) + Math.Pow(ring.Y + point.Y - centerY, 2));
        double outer = ring.Shape.Points.Max(Radius);
        double inner = ring.Shape.Points.Min(Radius);
        Assert.Equal(hole / 100d, inner / outer, 8);
        OfficePoint first = ring.Shape.Points[0];
        Assert.Equal(centerX + outer, ring.X + first.X, 8);
        Assert.Equal(centerY, ring.Y + first.Y, 8);
        OfficeDrawingText label = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "7");
        double labelRadius = Math.Sqrt(Math.Pow(label.X + label.Width / 2 - centerX, 2) + Math.Pow(label.Y + label.Height / 2 - centerY, 2));
        Assert.InRange(labelRadius, inner, outer);
        Assert.Equal((inner + outer) / 2, labelRadius, 8);
    }

    [Fact]
    public void RadialGeometry_FirstSeriesIsInnermostAndRingsMeetWithoutGaps() {
        var data = new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Inner", new[] { 1d }, null, color: OfficeColor.Parse("#FF0000")),
            new OfficeChartSeries("Outer", new[] { 1d }, null, color: OfficeColor.Parse("#0000FF"))
        });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Rings", null,
            OfficeChartKind.Doughnut, data, 400, 400, null, new OfficeChartLayout(showLegend: false),
            new OfficeChartRadialLayout(doughnutHolePercent: 50)));
        var rings = drawing.Shapes.Where(s => s.Shape.Kind == OfficeShapeKind.Polygon).ToArray();
        Assert.Equal(2, rings.Length);
        Assert.Equal(OfficeColor.Parse("#FF0000"), rings[0].Shape.FillColor);
        Assert.Equal(OfficeColor.Parse("#0000FF"), rings[1].Shape.FillColor);
        double[] Radii(OfficeDrawingShape ring) {
            double cx = ring.Shape.Width / 2, cy = ring.Shape.Height / 2;
            return ring.Shape.Points.Select(p => Math.Sqrt(Math.Pow(p.X - cx, 2) + Math.Pow(p.Y - cy, 2))).ToArray();
        }
        Assert.Equal(Radii(rings[0]).Max(), Radii(rings[1]).Min(), 8);
        Assert.Equal(0.5, Radii(rings[0]).Min() / Radii(rings[1]).Max(), 8);
    }

    [Fact]
    public void RadialGeometry_RejectsValuesOutsideNativeContract() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartRadialLayout(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartRadialLayout(361));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartRadialLayout(doughnutHolePercent: 9));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartRadialLayout(doughnutHolePercent: 91));
    }
}
