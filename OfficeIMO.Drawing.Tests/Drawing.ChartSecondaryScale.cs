using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartSecondaryScaleTests {
    [Fact]
    public void SecondaryOneSidedMaximumOutsideDataStillClipsItsLine() {
        OfficeColor primary = OfficeColor.Parse("#2A9D8F");
        OfficeColor secondary = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Primary", new[] { 1d, 2d }, null, primary),
            new OfficeSeries("Secondary", new[] { 10d, 20d }, null, secondary, null, true,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var layout = new OfficeChartLayout(showLegend: false, showMarkers: false)
            .WithSecondaryValueAxis(new OfficeChartValueAxisLayout(maximum: 5));
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartType.ColumnClustered, data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240), layout: layout));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                Assert.NotEqual(secondary, raster.GetPixel(x, y));
    }

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
    [InlineData(OfficeChartType.ColumnClustered)]
    [InlineData(OfficeChartType.BarClustered)]
    public void ValueAxisTicks_KeepInsideOutsideAndCrossPositions(OfficeChartType kind) {
        var data = new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeSeries("Values", new[] { 1d, 2d }) });
        bool horizontal = kind == OfficeChartType.BarClustered;
        double Position(OfficeChartAxisTickMark mark) {
            var layout = new OfficeChartLayout(showLegend: false,
                horizontalAxisMajorTickMark: horizontal ? mark : OfficeChartAxisTickMark.None,
                verticalAxisMajorTickMark: horizontal ? OfficeChartAxisTickMark.None : mark);
            var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, kind, data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240), layout: layout));
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
            new OfficeSeries("Primary", new[] { 100d, 200d }),
            new OfficeSeries("Secondary", new[] { 1d, 2d }, null, null, null, true,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        OfficeDrawing Render(OfficeChartAxisTickMark mark) {
            var layout = new OfficeChartLayout(showLegend: false).WithSecondaryValueAxis(
                new OfficeChartValueAxisLayout(minimum: 0, maximum: 4, majorUnit: 1, minorUnit: 0.5,
                    majorTickMark: mark, minorTickMark: mark));
            return OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
                OfficeChartType.ColumnClustered, data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240), layout: layout));
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
    [InlineData(OfficeChartType.ColumnClustered)]
    [InlineData(OfficeChartType.BarClustered)]
    public void SecondaryAxis_RendersAutomaticMinorTicksWhenVisible(OfficeChartType kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Primary", new[] { 100d, 200d }),
            new OfficeSeries("Secondary", new[] { 1d, 2d }, null, null, null, true,
                renderKind: kind, axisGroup: OfficeChartAxisGroup.Secondary) });
        int ShapeCount(OfficeChartAxisTickMark mark) {
            var layout = new OfficeChartLayout(showLegend: false).WithSecondaryValueAxis(
                new OfficeChartValueAxisLayout(minimum: 0, maximum: 4, majorUnit: 1,
                    minorTickMark: mark));
            return OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, kind,
                data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240), layout: layout)).Shapes.Count;
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
            new OfficeSeries("Primary", new[] { 100d, 200d }),
            new OfficeSeries("Secondary", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Scale", null,
            OfficeChartType.ColumnClustered, data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240), layout: layout));
        var labels = drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text).ToArray();
        Assert.Contains("200.0", labels);
        Assert.Contains("400%", labels);
        Assert.DoesNotContain("20000%", labels);
        Assert.Null(primaryLayout.SecondaryValueAxis);
        Assert.Null(OfficeChartLayout.Default.SecondaryValueAxis);
    }

    [Fact]
    public void MixedSecondarySeriesDoNotApplyPercentLabelsToAnOrdinaryAxis() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Primary", new[] { 5d, 6d }),
            new OfficeSeries("Share", new[] { 1d, 2d }, null, null, null, true,
                renderKind: OfficeChartType.LineStacked100, axisGroup: OfficeChartAxisGroup.Secondary),
            new OfficeSeries("Amount", new[] { 2d, 3d }, null, null, null, true,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var snapshot = new OfficeChartSnapshot("", null, OfficeChartType.ColumnClustered, data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240),
            layout: new OfficeChartLayout(showLegend: false).WithSecondaryValueAxis(new OfficeChartValueAxisLayout()));
        var labels = OfficeChartDrawingRenderer.Render(snapshot).Elements.OfType<OfficeDrawingText>()
            .Select(text => text.Text).ToArray();
        Assert.Contains("3", labels);
        Assert.DoesNotContain(labels, label => label.Contains('%'));
    }

    [Fact]
    public void LogarithmicSecondaryAxisSpacesPowersEvenlyAndLabelsSourceValues() {
        OfficeColor secondary = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeSeries("Primary", new[] { 1d, 2d, 3d }),
            new OfficeSeries("Secondary", new[] { 10d, 100d, 1000d }, null, secondary, null, false,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var layout = new OfficeChartLayout(showLegend: false)
            .WithSecondaryValueAxis(new OfficeChartValueAxisLayout().WithLogarithmicScale(10D));
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartType.ColumnClustered, data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240), layout: layout));

        // 10, 100 and 1000 are one power apart, so the line falls by the same height on each step.
        static IEnumerable<OfficeDrawingShape> Shapes(OfficeDrawing source) => source.Shapes.Concat(
            source.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => Shapes(group.Drawing)));
        double[] ys = Shapes(drawing).Where(shape => shape.Shape.StrokeColor == secondary && shape.Shape.Kind == OfficeShapeKind.Line)
            .SelectMany(shape => shape.Shape.Points.Select(point => Math.Round(shape.Y + point.Y, 6))).Distinct().OrderByDescending(y => y).ToArray();
        Assert.Equal(3, ys.Length);
        Assert.Equal(ys[0] - ys[1], ys[1] - ys[2], 3);
        string[] labels = drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text).ToArray();
        Assert.Contains("10", labels);
        Assert.Contains("100", labels);
    }

    [Fact]
    public void ReversedSecondaryAxisPutsItsMinimumAtTheTop() {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Primary", new[] { 1d, 2d }),
            new OfficeSeries("Secondary", new[] { 40d, 50d }, null, null, null, false,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var layout = new OfficeChartLayout(showLegend: false).WithSecondaryValueAxis(
            new OfficeChartValueAxisLayout(minimum: 0, maximum: 50, majorUnit: 10).WithReversedScale(true));
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartType.ColumnClustered, data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240), layout: layout));

        OfficeDrawingText[] texts = drawing.Elements.OfType<OfficeDrawingText>().ToArray();
        OfficeDrawingText zero = Assert.Single(texts, text => text.Text == "0" && text.X > 300);
        OfficeDrawingText fifty = Assert.Single(texts, text => text.Text == "50");
        Assert.True(zero.Y < fifty.Y);
    }

    // Bounds within one power of the base plot as less than one major unit; the axis keeps its bounds as its ticks.
    [Theory]
    [InlineData(1D, 2D, new[] { "1", "2" })]
    [InlineData(1D, 100D, new[] { "1", "10", "100" })]
    public void LogarithmicAxesDrawBoundsWithinOnePower(double minimum, double maximum, string[] expected) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Primary", new[] { 1.2d, 1.8d }),
            new OfficeSeries("Secondary", new[] { 1.5d, 1.9d }, null, null, null, true,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var secondary = new OfficeChartLayout(showLegend: false)
            .WithSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum, maximum).WithLogarithmicScale(10D));
        var primary = new OfficeChartLayout(showLegend: false, verticalAxisMinimum: minimum, verticalAxisMaximum: maximum)
            .WithLogarithmicValueAxes(null, 10D);
        foreach ((OfficeChartLayout layout, bool right) in new[] { (secondary, true), (primary, false) }) {
            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
                OfficeChartType.ColumnClustered, data, global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(240), layout: layout));
            // Only the labels beside the logarithmic axis: the secondary axis is on the right, the primary on the left.
            double[] labels = drawing.Elements.OfType<OfficeDrawingText>().Where(text => right ? text.X > 180D : text.X + text.Width < 180D)
                .Select(text => double.TryParse(text.Text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value) ? value : double.NaN)
                .Where(value => !double.IsNaN(value)).OrderBy(value => value).ToArray();
            Assert.Equal(expected.Select(label => double.Parse(label, System.Globalization.CultureInfo.InvariantCulture)), labels);
        }
    }
}
