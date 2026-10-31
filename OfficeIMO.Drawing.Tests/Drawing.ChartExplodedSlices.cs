using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartExplodedSlicesTests {
    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void ExplodedSlice_OffsetsOnlySelectedPointAndFitsCanvas(OfficeChartType kind) {
        var series = new OfficeSeries("Status", new[] { 5d, 5d })
            .WithPointExplosions(new[] { 25, 0 });
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { series });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot("", null, kind, data, global::OfficeIMO.OfficeLength.Points(320), global::OfficeIMO.OfficeLength.Points(240),
                null, new OfficeChartLayout(showLegend: false)), false);
        OfficeDrawingShape[] slices = drawing.Shapes
            .Where(item => item.Shape.Kind == OfficeShapeKind.Polygon).ToArray();
        Assert.Equal(2, slices.Length);
        double firstX = slices[0].X + slices[0].Shape.Points[0].X;
        double secondX = slices[1].X + slices[1].Shape.Points[0].X;
        if (kind == OfficeChartType.Pie) {
            double centerY = slices[0].Y + slices[0].Shape.Points[0].Y;
            OfficePoint outer = slices[0].Shape.Points[1];
            double radius = Math.Sqrt(Math.Pow(slices[0].X + outer.X - firstX, 2) +
                Math.Pow(slices[0].Y + outer.Y - centerY, 2));
            Assert.Equal(radius * 0.25D, firstX - secondX, 7);
        } else {
            Assert.True(firstX > secondX);
        }
        foreach (OfficeDrawingShape slice in slices)
            foreach (OfficePoint point in slice.Shape.Points) {
                Assert.InRange(slice.X + point.X, -0.000001D, 320.000001D);
                Assert.InRange(slice.Y + point.Y, -0.000001D, 240.000001D);
            }
    }

    [Fact]
    public void InnerDoughnutRing_ExplosionUsesItsOwnRadius() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Inner", new[] { 5d, 5d }).WithPointExplosions(new[] { 25, 0 }),
            new OfficeSeries("Outer", new[] { 5d, 5d })
        });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot("", null, OfficeChartType.Doughnut, data, global::OfficeIMO.OfficeLength.Points(320), global::OfficeIMO.OfficeLength.Points(240),
                null, new OfficeChartLayout(showLegend: false)), false);
        OfficeDrawingShape[] slices = drawing.Shapes
            .Where(item => item.Shape.Kind == OfficeShapeKind.Polygon).ToArray();
        Assert.Equal(4, slices.Length);
        double explodedFirstX = slices[0].X + slices[0].Shape.Points[0].X;
        double innerSecondX = slices[1].X + slices[1].Shape.Points[0].X;
        double outerFirstY = slices[2].Y + slices[2].Shape.Points[0].Y;
        double outerSecondY = slices[3].Y + slices[3].Shape.Points[0].Y;
        double centerY = (outerFirstY + outerSecondY) / 2D;
        double innerOuterRadius = slices[1].Y + slices[1].Shape.Points[0].Y - centerY;
        Assert.Equal(innerOuterRadius * 0.25D, explodedFirstX - innerSecondX, 7);
    }

    [Fact]
    public void ExplosionValues_AreAlignedBoundedAndRetainedAcrossSeriesCopies() {
        var series = new OfficeSeries("Status", new[] { 5d, 5d })
            .WithPointExplosions(new[] { 25, 0 })
            .WithPointStyles(new OfficeChartPointStyle?[] { null, null })
            .WithLegendVisibility(false);
        Assert.Equal(new[] { 25, 0 }, series.PointExplosions);
        Assert.Throws<ArgumentException>(() => series.WithPointExplosions(new[] { 25 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => series.WithPointExplosions(new[] { 401, 0 }));
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { series });
        Assert.Throws<NotSupportedException>(() => OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot("", null, OfficeChartType.Line, data, global::OfficeIMO.OfficeLength.Points(320), global::OfficeIMO.OfficeLength.Points(240))));
    }
}
