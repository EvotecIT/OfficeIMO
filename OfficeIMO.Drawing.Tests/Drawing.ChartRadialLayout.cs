using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartRadialLayoutTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie, false)]
    [InlineData(OfficeChartKind.Pie, true)]
    [InlineData(OfficeChartKind.Doughnut, false)]
    [InlineData(OfficeChartKind.Doughnut, true)]
    public void EmptyOutsideLabelSelectionKeepsCompactRadialGeometry(
        OfficeChartKind kind, bool filtered) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", new[] { 2d, 1d })
        });
        var plain = new OfficeChartSnapshot("", null, kind, data, 140, 120,
            layout: new OfficeChartLayout(showLegend: false));
        var emptyOutside = new OfficeChartSnapshot("", null, kind, data, 140, 120,
            layout: new OfficeChartLayout(showLegend: false, showDataLabels: true,
                showDataLabelCategoryNames: filtered,
                dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd,
                dataLabelSeriesIndexes: filtered ? Array.Empty<int>() : null));
        OfficeDrawing expected = OfficeChartDrawingRenderer.Render(plain);
        OfficeDrawing actual = OfficeChartDrawingRenderer.Render(emptyOutside);
        Assert.Equal(expected.Shapes.Select(shape => (shape.X, shape.Y,
                shape.Shape.Width, shape.Shape.Height)),
            actual.Shapes.Select(shape => (shape.X, shape.Y,
                shape.Shape.Width, shape.Shape.Height)));
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void FiniteExtremeSlicesKeepGeometryAndPercentLabels(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", new[] { 1e308, 1e308 })
        });
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelPercentages: true);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            kind, data, 360, 240, layout: layout));
        OfficeDrawingShape[] slices = drawing.Shapes.Where(shape =>
            shape.Shape.Kind == OfficeShapeKind.Polygon).ToArray();
        Assert.Equal(2, slices.Length);
        Assert.All(slices, slice => Assert.True(slice.Shape.Width > 50D &&
            slice.Shape.Height > 50D, "Each equal finite slice should cover half the radial plot."));
        Assert.Equal(2, drawing.Elements.OfType<OfficeDrawingText>()
            .Count(label => label.Text == "50%"));
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void OutsideLabelsUseLeadersAndDoNotOverlapOnEitherSide(OfficeChartKind kind) {
        string[] categories = Enumerable.Range(0, 12).Select(index => $"Cat{index}").ToArray();
        var data = new OfficeChartData(categories, new[] {
            new OfficeChartSeries("Values", Enumerable.Repeat(1d, categories.Length).ToArray())
        });
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelCategoryNames: true, dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            kind, data, 420, 300, layout: layout));
        OfficeDrawingText[] labels = drawing.Elements.OfType<OfficeDrawingText>()
            .Where(text => text.Text.StartsWith("Cat", StringComparison.Ordinal)).ToArray();
        Assert.Equal(categories.Length, labels.Length);
        Assert.Equal(categories.Length, drawing.Shapes.Count(shape =>
            shape.Shape.Kind == OfficeShapeKind.Line));
        foreach (OfficeDrawingText label in labels) {
            Assert.InRange(label.X, 0, drawing.Width - label.Width);
            Assert.InRange(label.Y, 0, drawing.Height - label.Height);
        }
        double centerX = drawing.Width / 2D;
        foreach (OfficeDrawingText[] side in new[] {
            labels.Where(label => label.X + label.Width / 2D < centerX).OrderBy(label => label.Y).ToArray(),
            labels.Where(label => label.X + label.Width / 2D >= centerX).OrderBy(label => label.Y).ToArray()
        })
            for (int index = 1; index < side.Length; index++)
                Assert.True(side[index].Y >= side[index - 1].Y + side[index - 1].Height + 1D);
    }

    [Fact]
    public void OutsideLabelsShrinkPieBeforeEnteringLabelGutters() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", new[] { 1d, 1d })
        });
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelCategoryNames: true, dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Pie, data, 200, 240, layout: layout));
        OfficeDrawingText[] labels = drawing.Elements.OfType<OfficeDrawingText>()
            .Where(text => text.Text is "A" or "B").ToArray();
        Assert.Equal(2, labels.Length);
        OfficeDrawingShape[] slices = drawing.Shapes.Where(shape => shape.Shape.Kind == OfficeShapeKind.Polygon).ToArray();
        double left = slices.SelectMany(slice => slice.Shape.Points.Select(point => slice.X + point.X)).Min();
        double right = slices.SelectMany(slice => slice.Shape.Points.Select(point => slice.X + point.X)).Max();
        Assert.Contains(labels, label => label.X + label.Width < left);
        Assert.Contains(labels, label => label.X > right);
    }

    [Fact]
    public void OutsideLabelsRejectCanvasThatCannotFitEveryLabel() {
        string[] categories = Enumerable.Range(0, 30).Select(index => $"Item{index}").ToArray();
        var data = new OfficeChartData(categories, new[] {
            new OfficeChartSeries("Values", Enumerable.Repeat(1d, categories.Length).ToArray())
        });
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelCategoryNames: true, dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd);
        Assert.Throws<NotSupportedException>(() => OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot("", null, OfficeChartKind.Pie, data, 300, 100, layout: layout)));
    }

    [Fact]
    public void MultiRingOutsideLabelsRequireLeadersToBeDisabled() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Inner", new[] { 1d, 2d }),
            new OfficeChartSeries("Outer", new[] { 2d, 1d })
        });
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelCategoryNames: true, dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd);
        var snapshot = new OfficeChartSnapshot("", null, OfficeChartKind.Doughnut, data, 420, 300,
            layout: layout);
        Assert.Throws<NotSupportedException>(() => OfficeChartDrawingRenderer.Render(snapshot));
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Doughnut, data, 420, 300,
            layout: layout.WithDataLabelLeaderLines(false)));
        Assert.Equal(4, drawing.Elements.OfType<OfficeDrawingText>()
            .Count(text => text.Text is "A" or "B"));
    }

    [Fact]
    public void MultiRingOutsideLabelsAllowLeadersWhenOnlyOuterRingIsLabeled() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Inner", new[] { 1d, 2d }),
            new OfficeChartSeries("Outer", new[] { 2d, 1d })
        });
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelCategoryNames: true, dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd,
            dataLabelSeriesIndexes: new[] { 1 });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Doughnut, data, 420, 300, layout: layout));
        Assert.Equal(2, drawing.Shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Line));
    }

    [Fact]
    public void OutsideLabelsUsePaintedChartBackgroundForContrast() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", new[] { 1d, 1d })
        });
        var style = new OfficeChartStyle(backgroundColor: OfficeColor.White,
            plotAreaBackgroundColor: OfficeColor.Black);
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelCategoryNames: true, dataLabelPosition: OfficeChartDataLabelPosition.OutsideEnd);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Pie, data, 420, 300, style, layout));
        Assert.All(drawing.Elements.OfType<OfficeDrawingText>()
            .Where(text => text.Text is "A" or "B"), text => Assert.Equal(OfficeColor.Black, text.Color));
    }

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
