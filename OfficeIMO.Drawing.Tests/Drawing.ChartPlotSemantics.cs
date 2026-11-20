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
