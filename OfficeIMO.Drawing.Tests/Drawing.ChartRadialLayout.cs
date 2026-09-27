using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartRadialLayoutTests {
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
    public void RadialGeometry_RejectsValuesOutsideNativeContract() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartRadialLayout(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartRadialLayout(361));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartRadialLayout(doughnutHolePercent: 9));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartRadialLayout(doughnutHolePercent: 91));
    }
}
