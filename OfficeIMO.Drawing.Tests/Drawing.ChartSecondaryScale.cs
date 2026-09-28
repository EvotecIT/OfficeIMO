using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartSecondaryScaleTests {
    [Fact]
    public void ValueAxisLayout_RejectsInvalidNumericAndTickSettings() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartValueAxisLayout(minimum: double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartValueAxisLayout(maximum: double.PositiveInfinity));
        Assert.Throws<ArgumentException>(() => new OfficeChartValueAxisLayout(minimum: 2, maximum: 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartValueAxisLayout(majorUnit: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartValueAxisLayout(minorUnit: -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartValueAxisLayout(majorTickMark: (OfficeChartAxisTickMark)99));
    }

    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.BarClustered)]
    public void ValueAxisTicks_KeepInsideOutsideAndCrossPositions(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }) });
        bool horizontal = kind == OfficeChartKind.BarClustered;
        double Position(OfficeChartAxisTickMark mark) {
            var layout = new OfficeChartLayout(showLegend: false,
                horizontalAxisMajorTickMark: horizontal ? mark : OfficeChartAxisTickMark.None,
                verticalAxisMajorTickMark: horizontal ? OfficeChartAxisTickMark.None : mark);
            var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, kind, data, 360, 240, layout: layout));
            var ticks = drawing.Shapes.Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                (horizontal ? shape.Shape.Width == 0 && shape.Shape.Height == 4 : shape.Shape.Width == 4 && shape.Shape.Height == 0)).ToArray();
            Assert.NotEmpty(ticks);
            return horizontal ? ticks[0].Y : ticks[0].X;
        }
        double inside = Position(OfficeChartAxisTickMark.Inside);
        double outside = Position(OfficeChartAxisTickMark.Outside);
        double cross = Position(OfficeChartAxisTickMark.Cross);
        Assert.Equal(horizontal ? 4 : -4, outside - inside, 8);
        Assert.Equal((inside + outside) / 2, cross, 8);
    }

    [Fact]
    public void SecondaryAxis_RendersItsOwnMajorAndMinorTicks() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Primary", new[] { 100d, 200d }),
            new OfficeChartSeries("Secondary", new[] { 1d, 2d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        OfficeDrawing Render(OfficeChartAxisTickMark mark) {
            var layout = new OfficeChartLayout(showLegend: false).WithSecondaryValueAxis(
                new OfficeChartValueAxisLayout(minimum: 0, maximum: 4, majorUnit: 1, minorUnit: 0.5,
                    majorTickMark: mark, minorTickMark: mark));
            return OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
                OfficeChartKind.ColumnClustered, data, 360, 240, layout: layout));
        }
        var without = Render(OfficeChartAxisTickMark.None);
        var with = Render(OfficeChartAxisTickMark.Outside);
        Assert.Equal(without.Shapes.Count + 9, with.Shapes.Count);
        var ticks = with.Shapes.Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
            shape.Shape.Width == 4 && shape.Shape.Height == 0).ToArray();
        Assert.Equal(9, ticks.Length);
        Assert.Single(ticks.Select(tick => tick.X).Distinct());
    }

    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.BarClustered)]
    public void SecondaryAxis_RendersAutomaticMinorTicksWhenVisible(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Primary", new[] { 100d, 200d }),
            new OfficeChartSeries("Secondary", new[] { 1d, 2d }, null, null, null, true,
                renderKind: kind, axisGroup: OfficeChartAxisGroup.Secondary) });
        int ShapeCount(OfficeChartAxisTickMark mark) {
            var layout = new OfficeChartLayout(showLegend: false).WithSecondaryValueAxis(
                new OfficeChartValueAxisLayout(minimum: 0, maximum: 4, majorUnit: 1,
                    minorTickMark: mark));
            return OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, kind,
                data, 360, 240, layout: layout)).Shapes.Count;
        }
        Assert.Equal(ShapeCount(OfficeChartAxisTickMark.None) + 16,
            ShapeCount(OfficeChartAxisTickMark.Outside));
    }

    [Fact]
    public void SecondaryScaleAndFormatRemainIndependentOfPrimaryAxis() {
        var primaryLayout = new OfficeChartLayout(showLegend: false, verticalAxisMinimum: 0,
            verticalAxisMaximum: 200, verticalAxisNumberFormat: "0.0");
        var layout = primaryLayout.WithSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0,
            maximum: 4, majorUnit: 1, numberFormat: "0%"));
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Primary", new[] { 100d, 200d }),
            new OfficeChartSeries("Secondary", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Scale", null,
            OfficeChartKind.ColumnClustered, data, 360, 240, layout: layout));
        var labels = drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text).ToArray();
        Assert.Contains("200.0", labels);
        Assert.Contains("400%", labels);
        Assert.DoesNotContain("20000%", labels);
        Assert.Null(primaryLayout.SecondaryValueAxis);
        Assert.Null(OfficeChartLayout.Default.SecondaryValueAxis);
    }
}
