using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Pins how the shared chart renderer draws missing values, axis crossings and their tick labels.</summary>
public sealed class DrawingChartPlotSemanticsTests {
    private static readonly OfficeColor Ink = OfficeColor.Parse("#D900AA");
    private static readonly string[] Categories = { "Jan", "Feb", "Mar", "Apr", "May" };
    private static readonly double[] WithMissingValue = { 90, 95, double.NaN, 101, 99 };
    private static readonly double[] WithNegativeValues = { 30, -20, 45, -35, 10 };

    [Theory]
    [InlineData(OfficeChartType.Line, OfficeChartBlankDisplay.Gap, 2)]
    [InlineData(OfficeChartType.Line, OfficeChartBlankDisplay.Span, 3)]
    [InlineData(OfficeChartType.Line, OfficeChartBlankDisplay.Zero, 4)]
    [InlineData(OfficeChartType.Scatter, OfficeChartBlankDisplay.Gap, 2)]
    [InlineData(OfficeChartType.Scatter, OfficeChartBlankDisplay.Span, 3)]
    [InlineData(OfficeChartType.Scatter, OfficeChartBlankDisplay.Zero, 4)]
    public void MissingLineAndScatterValuesFollowTheBlankDisplay(OfficeChartType kind, OfficeChartBlankDisplay blanks, int segments) {
        OfficeDrawing drawing = Render(kind, new OfficeChartLayout(showLegend: false).WithBlankDisplay(blanks),
            Series(WithMissingValue, kind == OfficeChartType.Scatter));

        // Gap drops both segments that touch the missing value, Span replaces them with one, and Zero keeps both.
        Assert.Equal(segments, SeriesInk(drawing).Count(shape => shape.Shape.Kind == OfficeShapeKind.Line));
    }

    [Theory]
    [InlineData(OfficeChartBlankDisplay.Gap)]
    [InlineData(OfficeChartBlankDisplay.Span)]
    [InlineData(OfficeChartBlankDisplay.Zero)]
    public void MissingRadarValuesFollowTheBlankDisplay(OfficeChartBlankDisplay blanks) {
        OfficeDrawing drawing = Render(OfficeChartType.Radar, new OfficeChartLayout(showLegend: false).WithBlankDisplay(blanks),
            Series(WithMissingValue));

        OfficeDrawingShape[] ink = Shapes(drawing).Where(shape => shape.Shape.StrokeColor == Ink).ToArray();
        if (blanks == OfficeChartBlankDisplay.Gap) {
            // Only neighbouring values are joined: Jan-Feb, Apr-May and May back to Jan.
            Assert.Equal(3, ink.Length);
            Assert.All(ink, shape => Assert.Equal(OfficeShapeKind.Line, shape.Shape.Kind));
        } else {
            OfficeDrawingShape outline = Assert.Single(ink);
            Assert.Equal(blanks == OfficeChartBlankDisplay.Span ? 4 : 5, outline.Shape.Points.Count);
        }
    }

    [Theory]
    [InlineData(OfficeChartBlankDisplay.Gap)]
    [InlineData(OfficeChartBlankDisplay.Span)]
    public void AreaSeriesDrawMissingValuesAtZero(OfficeChartBlankDisplay blanks) {
        OfficeDrawing drawing = Render(OfficeChartType.Area, new OfficeChartLayout(showLegend: false).WithBlankDisplay(blanks),
            Series(WithMissingValue));

        Assert.Single(Shapes(drawing), shape => shape.Shape.FillColor == Ink);
    }

    [Theory]
    [InlineData(OfficeChartAxisTickLabelPosition.NextTo)]
    [InlineData(OfficeChartAxisTickLabelPosition.Low)]
    [InlineData(OfficeChartAxisTickLabelPosition.High)]
    public void CategoryLabelsFollowTickLabelPositionAroundAZeroCrossing(OfficeChartAxisTickLabelPosition position) {
        OfficeDrawing drawing = Render(OfficeChartType.ColumnClustered,
            new OfficeChartLayout(showLegend: false, horizontalAxisTickLabelPosition: position), Series(WithNegativeValues));

        double labelTop = Text(drawing, "Jan").Y;
        double zeroY = TextCenterY(drawing, "0");
        if (position == OfficeChartAxisTickLabelPosition.NextTo) {
            // The category axis crosses at zero, so its labels sit just below the zero line inside the plot.
            Assert.InRange(labelTop, zeroY, zeroY + 15D);
        } else if (position == OfficeChartAxisTickLabelPosition.Low) {
            Assert.True(labelTop > TextCenterY(drawing, "-40"), "Low labels belong below the plot.");
        } else {
            Assert.True(labelTop < TextCenterY(drawing, "60"), "High labels belong above the plot.");
        }
    }

    [Theory]
    [InlineData(OfficeChartAxisTickLabelPosition.NextTo)]
    [InlineData(OfficeChartAxisTickLabelPosition.Low)]
    public void HorizontalBarCategoryLabelsFollowTickLabelPositionAroundAZeroCrossing(OfficeChartAxisTickLabelPosition position) {
        OfficeDrawing drawing = Render(OfficeChartType.BarClustered,
            new OfficeChartLayout(showLegend: false, verticalAxisTickLabelPosition: position), Series(WithNegativeValues));

        OfficeDrawingText label = Text(drawing, "Jan");
        OfficeDrawingText zero = Text(drawing, "0");
        double zeroX = zero.X + zero.Width / 2D;
        if (position == OfficeChartAxisTickLabelPosition.NextTo) {
            Assert.InRange(label.X + label.Width, zeroX - 12D, zeroX);
        } else {
            OfficeDrawingText minimum = Text(drawing, "-40");
            Assert.True(label.X + label.Width < minimum.X + minimum.Width / 2D, "Low labels belong left of the plot.");
        }
    }

    [Fact]
    public void ClusteredColumnsGrowFromAnExplicitAxisCrossing() {
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false).WithAxisCrossings(
            OfficeChartAxisCrossingPosition.AutoZero, -10D, OfficeChartAxisCrossingPosition.AutoZero, null);
        OfficeDrawing drawing = Render(OfficeChartType.ColumnClustered, layout, Series(WithNegativeValues));

        double minusTenY = (TextCenterY(drawing, "0") + TextCenterY(drawing, "-20")) / 2D;
        OfficeDrawingShape[] bars = SeriesFill(drawing).ToArray();
        Assert.Equal(WithNegativeValues.Length, bars.Length);
        // Positive columns end at -10 below; negative columns start there and hang down.
        Assert.All(bars, bar => Assert.True(
            Math.Abs(bar.Y + bar.Shape.Height - minusTenY) < 0.5D || Math.Abs(bar.Y - minusTenY) < 0.5D,
            $"The column at {bar.Y:0.##}-{bar.Y + bar.Shape.Height:0.##} does not start at the crossing {minusTenY:0.##}."));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataLabelsBeyondAnExplicitMaximumFollowTheOverMaximumSetting(bool showOverMaximum) {
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true,
            verticalAxisMinimum: 0D, verticalAxisMaximum: 30D).WithDataLabelsOverMaximum(showOverMaximum);
        OfficeDrawing drawing = Render(OfficeChartType.ColumnClustered, layout, Series(new double[] { 45, 20, 10, 25, 5 }));

        Assert.Contains(Texts(drawing), text => text.Text == "20");
        Assert.Equal(showOverMaximum, Texts(drawing).Any(text => text.Text == "45"));
    }

    [Fact]
    public void TopRightLegendStartsAtThePlotTop() {
        OfficeSeries[] series = { new OfficeSeries("A", new double[] { 1, 2, 3, 4, 5 }), new OfficeSeries("B", new double[] { 2, 3, 4, 5, 6 }) };
        OfficeDrawing right = Render(OfficeChartType.ColumnClustered,
            new OfficeChartLayout(legendPosition: OfficeChartLegendPosition.Right), series);
        OfficeDrawing topRight = Render(OfficeChartType.ColumnClustered,
            new OfficeChartLayout(legendPosition: OfficeChartLegendPosition.TopRight), series);

        Assert.True(Text(topRight, "A").Y < TextCenterY(topRight, "4"), "A top-right legend starts beside the top of the plot.");
        Assert.True(Text(right, "A").Y > Text(topRight, "A").Y + 50D, "A right legend stays centred on the plot.");
        Assert.Equal(Text(right, "A").X, Text(topRight, "A").X, 3);
    }

    [Fact]
    public void AreaLabelsSkipEmptyCellsButKeepRealZeros() {
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true);
        OfficeDrawing drawing = Render(OfficeChartType.Area, layout, Series(new double[] { 7, double.NaN, 9, 0, 11 }));

        // One "0" is the value-axis label; the empty cell adds none, the real zero adds one.
        Assert.Equal(2, Texts(drawing).Count(text => text.Text == "0"));
    }

    [Fact]
    public void InsideEndLabelsFollowAColumnGrowingFromAnAxisCrossing() {
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true,
            dataLabelPosition: OfficeChartDataLabelPosition.InsideEnd).WithAxisCrossings(
            OfficeChartAxisCrossingPosition.AutoZero, -10D, OfficeChartAxisCrossingPosition.AutoZero, null);
        OfficeDrawing drawing = Render(OfficeChartType.ColumnClustered, layout, Series(new double[] { 30, -5, 45, -35, 10 }));

        // -5 is above the -10 crossing, so its column grows up and the label sits inside its top.
        OfficeDrawingShape column = SeriesFill(drawing).OrderBy(bar => bar.X).ElementAt(1);
        OfficeDrawingText label = Text(drawing, "-5");
        Assert.InRange(label.Y, column.Y, column.Y + column.Shape.Height / 2D);
    }

    [Fact]
    public void RadarSpanJoinsTheTwoRemainingPoints() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, OfficeChartType.Radar,
            new OfficeChartData(new[] { "N", "E", "S", "W" }, new[] { Series(new double[] { 10, double.NaN, 20, double.NaN }) }),
            global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(300),
            layout: new OfficeChartLayout(showLegend: false).WithBlankDisplay(OfficeChartBlankDisplay.Span)));

        OfficeDrawingShape line = Assert.Single(Shapes(drawing), shape => shape.Shape.StrokeColor == Ink);
        Assert.Equal(OfficeShapeKind.Line, line.Shape.Kind);
    }

    [Fact]
    public void PercentStackedLabelsCompareTheirStackTopWithTheMaximum() {
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true,
            verticalAxisMinimum: 0D, verticalAxisMaximum: 0.5D).WithDataLabelsOverMaximum(false);
        OfficeDrawing drawing = Render(OfficeChartType.ColumnStacked100, layout,
            new OfficeSeries("Low", new double[] { 30, 30, 30, 30, 30 }), new OfficeSeries("High", new double[] { 70, 70, 70, 70, 70 }));

        // The low segment tops out at 30%, below the 50% maximum; the high one reaches 100%.
        Assert.Equal(5, Texts(drawing).Count(text => text.Text == "30"));
        Assert.DoesNotContain(Texts(drawing), text => text.Text == "70");
    }

    [Fact]
    public void ZeroStackMembersAreJudgedAtTheirStackTop() {
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true,
            verticalAxisMinimum: 0D, verticalAxisMaximum: 30D).WithDataLabelsOverMaximum(false);
        OfficeDrawing drawing = Render(OfficeChartType.ColumnStacked, layout,
            new OfficeSeries("Low", new double[] { 40, 40, 40, 40, 10 }), new OfficeSeries("None", new double[] { 0, 0, 0, 0, 0 }));

        // Four zeros sit on stacks topping out at 40, beyond the maximum; the last sits at 10. The other "0" is the axis label.
        Assert.Equal(2, Texts(drawing).Count(text => text.Text == "0"));
    }

    [Fact]
    public void RadarAxisLabelsUseDisplayUnits() {
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false).WithAxisDisplayUnits(null, null, 1000D, "Thousands");
        OfficeDrawing drawing = Render(OfficeChartType.Radar, layout, Series(new double[] { 1000, 2000, 3000, 4000, 4000 }));

        Assert.Contains(Texts(drawing), text => text.Text == "Thousands");
        Assert.DoesNotContain(Texts(drawing), text => text.Text == "4000" || text.Text == "4,000");
    }

    [Fact]
    public void LogarithmicValueAxisSpacesPowersEvenlyAndSkipsNonPositiveValues() {
        OfficeDrawing drawing = Render(OfficeChartType.ColumnClustered,
            new OfficeChartLayout(showLegend: false).WithLogarithmicValueAxes(null, 10D), Series(new double[] { 10, 100, 1000, 0, 1 }));

        double[] heights = SeriesFill(drawing).OrderBy(bar => bar.X).Select(bar => bar.Shape.Height).ToArray();
        // 10, 100 and 1000 rise one, two and three powers above 1; 0 has no bar and 1 has no height.
        Assert.Equal(3, heights.Length);
        Assert.Equal(2D, heights[1] / heights[0], 2);
        Assert.Equal(3D, heights[2] / heights[0], 2);
        Assert.Contains(Texts(drawing), text => text.Text == "100");
    }

    [Theory]
    [InlineData(OfficeChartType.ColumnClustered)]
    [InlineData(OfficeChartType.BarClustered)]
    public void ReversedValueAxisPutsTheMinimumAtTheFarEnd(OfficeChartType kind) {
        bool horizontalBars = kind == OfficeChartType.BarClustered;
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false)
            .WithReversedValueAxes(horizontal: horizontalBars, vertical: !horizontalBars);
        OfficeDrawing drawing = Render(kind, layout, Series(new double[] { 10, 20, 30, 40, 50 }));

        OfficeDrawingText zero = Text(drawing, "0");
        OfficeDrawingText fifty = Text(drawing, "50");
        OfficeDrawingShape longest = SeriesFill(drawing).OrderByDescending(bar => horizontalBars ? bar.Shape.Width : bar.Shape.Height).First();
        if (horizontalBars) {
            // Zero is on the right and bars grow leftwards from it.
            Assert.True(zero.X > fifty.X);
            Assert.InRange(longest.X + longest.Shape.Width, zero.X, zero.X + zero.Width);
        } else {
            // Zero is at the top and columns hang down from it.
            Assert.True(zero.Y < fifty.Y);
            Assert.InRange(longest.Y, zero.Y, zero.Y + zero.Height);
        }
    }

    [Fact]
    public void ManualLayoutsPlaceThePlotTitleAndLegend() {
        var plot = new OfficeChartManualLayout(0.3D, 0.3D, 0.4D, 0.4D, OfficeChartLayoutMode.Edge, OfficeChartLayoutMode.Edge,
            target: OfficeChartLayoutTarget.Inner);
        var title = new OfficeChartManualLayout(0.05D, 0.85D, xMode: OfficeChartLayoutMode.Edge, yMode: OfficeChartLayoutMode.Edge);
        var legend = new OfficeChartManualLayout(0.8D, 0.05D, 0.15D, 0.2D, OfficeChartLayoutMode.Edge, OfficeChartLayoutMode.Edge);
        OfficeChartLayout layout = new OfficeChartLayout().WithManualLayouts(plot, title, legend);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", "Placed", OfficeChartType.ColumnClustered,
            new OfficeChartData(Categories, new[] { Series(new double[] { 1, 2, 3, 4, 5 }) }),
            global::OfficeIMO.OfficeLength.Points(500), global::OfficeIMO.OfficeLength.Points(300), layout: layout));

        Assert.All(SeriesFill(drawing).Where(bar => bar.Shape.Width > 12D), bar => {
            Assert.InRange(bar.X, 150D, 350D);
            Assert.InRange(bar.Y + bar.Shape.Height, 90D, 210.5D);
        });
        OfficeDrawingText titleText = Text(drawing, "Placed");
        Assert.InRange(titleText.X, 24D, 26D);
        Assert.InRange(titleText.Y, 254D, 256D);
        OfficeDrawingText legendText = Text(drawing, "Sales");
        Assert.InRange(legendText.X, 400D, 430D);
        Assert.InRange(legendText.Y, 15D, 60D);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void VaryColorsByPointColoursASingleSeriesAndListsCategories(int seriesCount) {
        OfficeSeries[] series = Enumerable.Range(0, seriesCount)
            .Select(index => new OfficeSeries("S" + index, new double[] { 1, 2, 3, 4, 5 })).ToArray();
        OfficeDrawing drawing = Render(OfficeChartType.ColumnClustered,
            new OfficeChartLayout().WithVaryColorsByPoint(true), series);

        int distinctFills = Shapes(drawing).Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.Height > 0D &&
            shape.Shape.Width > 12D && shape.Shape.Width < 60D && shape.Shape.FillColor.HasValue).Select(shape => shape.Shape.FillColor).Distinct().Count();
        if (seriesCount == 1) {
            Assert.True(distinctFills >= Categories.Length, $"Only {distinctFills} colours for {Categories.Length} points.");
            Assert.Equal(2, Texts(drawing).Count(text => text.Text == "Mar"));
        } else {
            Assert.DoesNotContain(Texts(drawing), text => text.Text == "Mar" && text.X > 400D);
            Assert.Contains(Texts(drawing), text => text.Text == "S1");
        }
    }

    [Theory]
    [InlineData(OfficeChartType.ColumnClustered)]
    [InlineData(OfficeChartType.BarClustered)]
    public void BarsOnTickMarksCentreOnTheCategoryTicksAndClipAtThePlotEdge(OfficeChartType kind) {
        bool horizontalBars = kind == OfficeChartType.BarClustered;
        OfficeDrawing drawing = Render(kind, new OfficeChartLayout(showLegend: false).WithCategoryPointsOnTickMarks(true),
            Series(new double[] { 10, 20, 30, 40, 50 }));

        OfficeDrawingShape[] bars = SeriesFill(drawing).ToArray();
        double Center(OfficeDrawingShape bar) => horizontalBars ? bar.Y + bar.Shape.Height / 2D : bar.X + bar.Shape.Width / 2D;
        double LabelCenter(string name) {
            OfficeDrawingText label = Text(drawing, name);
            return horizontalBars ? label.Y + label.Height / 2D : label.X + label.Width / 2D;
        }
        // The first and last categories sit on the plot edges, so their bars are half outside and clipped;
        // the others are centred on their tick marks.
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingGroup>(), group => SeriesFill(group.Drawing).Count() == bars.Length);
        Assert.Contains(bars, bar => Math.Abs(Center(bar) - LabelCenter("Jan")) < 1D);
        Assert.Contains(bars, bar => Math.Abs(Center(bar) - LabelCenter("Apr")) < 1D);
    }

    [Fact]
    public void StackedAndScatterLabelsBeyondAnExplicitMaximumAreHidden() {
        OfficeChartLayout stackedLayout = new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true,
            verticalAxisMinimum: 0D, verticalAxisMaximum: 30D).WithDataLabelsOverMaximum(false);
        OfficeDrawing stacked = Render(OfficeChartType.ColumnStacked, stackedLayout,
            new OfficeSeries("Low", new double[] { 11, 12, 13, 14, 15 }), new OfficeSeries("High", new double[] { 21, 22, 3, 4, 5 }));
        // 13 + 3 and the other short stacks stay below 30; 11 + 21 and 12 + 22 go beyond it.
        Assert.DoesNotContain(Texts(stacked), text => text.Text is "21" or "22");
        Assert.Contains(Texts(stacked), text => text.Text == "3");
        Assert.Contains(Texts(stacked), text => text.Text == "11");

        OfficeChartLayout scatterLayout = new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true,
            horizontalAxisMinimum: 0D, horizontalAxisMaximum: 3D).WithDataLabelsOverMaximum(false);
        OfficeDrawing scatter = Render(OfficeChartType.Scatter, scatterLayout,
            new OfficeSeries("Points", new double[] { 7.3, 8.3, 9.3, 6.3, 5.3 }, new double[] { 1, 2, 3, 4, 5 }));
        Assert.Contains(Texts(scatter), text => text.Text == "9.3");
        Assert.DoesNotContain(Texts(scatter), text => text.Text is "6.3" or "5.3");
    }

    private static OfficeSeries Series(double[] values, bool numericX = false) =>
        new OfficeSeries("Sales", values, numericX ? values.Select((_, index) => index + 1D) : null, Ink, null, showMarkers: false);

    private static OfficeDrawing Render(OfficeChartType kind, OfficeChartLayout layout, params OfficeSeries[] series) =>
        OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, kind, new OfficeChartData(Categories, series),
            global::OfficeIMO.OfficeLength.Points(480), global::OfficeIMO.OfficeLength.Points(300), layout: layout));

    private static IEnumerable<OfficeDrawingShape> SeriesInk(OfficeDrawing drawing) =>
        Shapes(drawing).Where(shape => shape.Shape.StrokeColor == Ink && shape.Shape.FillColor == null);

    private static IEnumerable<OfficeDrawingShape> SeriesFill(OfficeDrawing drawing) =>
        Shapes(drawing).Where(shape => shape.Shape.FillColor == Ink);

    private static IEnumerable<OfficeDrawingShape> Shapes(OfficeDrawing drawing) => drawing.Shapes.Concat(
        drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => Shapes(group.Drawing)));

    private static IEnumerable<OfficeDrawingText> Texts(OfficeDrawing drawing) => drawing.Elements.OfType<OfficeDrawingText>().Concat(
        drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => Texts(group.Drawing)));

    private static OfficeDrawingText Text(OfficeDrawing drawing, string text) => Assert.Single(Texts(drawing), item => item.Text == text);

    private static double TextCenterY(OfficeDrawing drawing, string text) {
        OfficeDrawingText item = Text(drawing, text);
        return item.Y + item.Height / 2D;
    }
}
