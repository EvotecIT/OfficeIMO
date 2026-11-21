using System.Collections;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartErrorBarsAndForecastsTests {
    private static readonly OfficeColor Ink = OfficeColor.Parse("#1F77B4");
    private static readonly OfficeColor Trend = OfficeColor.Parse("#00A0D9");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StatisticalErrorBarsReadTheSeriesOnceForAllPoints(bool standardError) {
        OfficeChartErrorBarValueType kind = standardError ? OfficeChartErrorBarValueType.StandardError : OfficeChartErrorBarValueType.StandardDeviation;
        var values = new CountingList(Enumerable.Range(1, 2_000).Select(value => (double)value).ToArray());
        var bars = new OfficeChartErrorBars(kind, 1D, OfficeChartErrorBarDirection.Both, endCaps: true, alongCategoryAxis: false);

        for (int index = 0; index < values.Count; index++)
            Assert.True(bars.TryGetSpan(values, index, values[index], out _, out _, out _));

        // Indexing each point once plus one pass for the statistics; per-point statistics would read millions.
        Assert.True(values.Reads <= 3 * values.Count, $"Read {values.Reads} values for {values.Count} points.");
    }

    [Fact]
    public void StatisticsFollowADifferentValueList() {
        var bars = new OfficeChartErrorBars(OfficeChartErrorBarValueType.StandardError, 1D, OfficeChartErrorBarDirection.Both, true, false);

        Assert.True(bars.TryGetSpan(new double[] { 1, 3 }, 0, 1, out _, out double firstLow, out double firstHigh));
        Assert.True(bars.TryGetSpan(new double[] { 1, 5 }, 0, 1, out _, out double secondLow, out double secondHigh));

        Assert.Equal(2D, firstHigh - firstLow, 10);
        Assert.Equal(4D, secondHigh - secondLow, 10);
    }

    [Fact]
    public void ForecastsAtOrBeyondOneBillionAreKept() {
        OfficeChartTrendline trendline = new OfficeChartTrendline(OfficeChartTrendlineType.Linear).WithNativeDetails(1E9D, 5E9D, null);

        Assert.Equal(1E9D, trendline.Forward);
        Assert.Equal(5E9D, trendline.Backward);
    }

    [Fact]
    public void ScatterForecastsWidenTheAxisOnlyFromPlottedPoints() {
        OfficeSeries series = new OfficeSeries("Sales", new[] { 1D, 2D, double.NaN }, new[] { 1D, 2D, 1000D }, Ink, null, showMarkers: true)
            .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Linear, color: Trend).WithNativeDetails(1D, 0D, null));

        OfficeDrawing drawing = Render(OfficeChartType.Scatter, new OfficeChartLayout(showLegend: false), new[] { "1", "2", "3" }, series);

        // The x = 1000 point has no y value, so the axis covers 1 to 3 (the forecast), not 1001.
        double largest = Texts(drawing).Select(text => double.TryParse(text.Text, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out double value) ? value : 0D).Max();
        Assert.True(largest < 10D, $"The axis reaches {largest}.");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ForecastsOfATrendlineThatCannotBeDrawnDoNotWidenTheAxis(bool polynomial) {
        // One valid pair cannot carry a linear fit, and three cannot carry a cubic one, so nothing is drawn.
        OfficeSeries series = polynomial
            ? new OfficeSeries("Sales", new[] { 1D, 4D, 2D }, new[] { 1D, 2D, 3D }, Ink, null, showMarkers: true)
                .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Polynomial, order: 3, color: Trend).WithNativeDetails(1000D, 0D, null))
            : new OfficeSeries("Sales", new[] { 1D, double.NaN }, new[] { 1D, 2D }, Ink, null, showMarkers: true)
                .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Linear, color: Trend).WithNativeDetails(1000D, 0D, null));

        OfficeDrawing drawing = Render(OfficeChartType.Scatter, new OfficeChartLayout(showLegend: false), new[] { "1", "2", "3" }, series);

        Assert.Empty(TrendPoints(drawing));
        double largest = Texts(drawing).Select(text => double.TryParse(text.Text, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out double value) ? value : 0D).Max();
        Assert.True(largest < 10D, $"The axis reaches {largest}.");
    }

    [Fact]
    public void ForecastsOfATrendlineALogarithmicValueAxisCannotDrawDoNotWidenTheAxis() {
        // The linear fit of these values is about -49.5 everywhere, which a logarithmic y axis cannot draw.
        OfficeSeries series = new OfficeSeries("Sales", new[] { -100D, 1D, 1D, -100D }, new[] { 1D, 2D, 3D, 4D }, Ink, null, showMarkers: true)
            .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Linear, color: Trend).WithNativeDetails(1000D, 0D, null));

        OfficeDrawing drawing = Render(OfficeChartType.Scatter, new OfficeChartLayout(showLegend: false).WithLogarithmicValueAxes(10D, 10D),
            new[] { "1", "2", "3", "4" }, series);

        Assert.Empty(TrendPoints(drawing));
        double largest = Texts(drawing).Select(text => double.TryParse(text.Text, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out double value) ? value : 0D).Max();
        Assert.True(largest < 100D, $"The axis reaches {largest}.");
    }

    [Fact]
    public void BackwardForecastOnALogarithmicXAxisKeepsItsDrawablePart() {
        OfficeDrawing Draw(double backward) => Render(OfficeChartType.Scatter,
            new OfficeChartLayout(showLegend: false, horizontalAxisMinimum: 0.1D, horizontalAxisMaximum: 100D).WithLogarithmicValueAxes(10D, null),
            new[] { "1", "2" },
            new OfficeSeries("Sales", new[] { 1D, 2D }, new[] { 1D, 10D }, Ink, null, showMarkers: true)
                .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Linear, color: Trend).WithNativeDetails(0D, backward, null)));

        double withoutForecast = TrendPoints(Draw(0D)).Min(point => point.X);
        double withForecast = TrendPoints(Draw(2D)).Min(point => point.X);

        // A forecast back to x = -1 cannot be drawn on a log axis, but its part from 0.1 to 1 can.
        Assert.True(withForecast < withoutForecast - 50D, $"The forecast starts at {withForecast}, the data at {withoutForecast}.");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ScatterForecastsAreDrawnOnReversedXAxes(bool logarithmic, bool backward) {
        OfficeDrawing Draw(double forecast) {
            OfficeChartLayout layout = new OfficeChartLayout(showLegend: false).WithReversedValueAxes(true, false);
            if (logarithmic) layout = layout.WithLogarithmicValueAxes(10D, null);
            OfficeChartTrendline trendline = new OfficeChartTrendline(OfficeChartTrendlineType.Linear, color: Trend)
                .WithNativeDetails(backward ? 0D : forecast, backward ? forecast : 0D, null);
            return Render(OfficeChartType.Scatter, layout, new[] { "1", "2", "3" },
                new OfficeSeries("Sales", new[] { 1D, 2D, 3D }, new[] { 2D, 4D, 8D }, Ink, null, showMarkers: true).WithTrendline(trendline));
        }

        static double Span(OfficeDrawing drawing) {
            OfficePoint[] points = TrendPoints(drawing).ToArray();
            return points.Max(point => point.X) - points.Min(point => point.X);
        }

        // A forecast of one data unit either way stays on screen once the axis widens to show it, so the drawn
        // trendline covers more of the plot than the data alone.
        double data = Span(Draw(0D));
        double forecast = Span(Draw(1D));
        Assert.True(forecast > data + 10D, $"The trendline spans {forecast} with the forecast and {data} without it.");
    }

    [Fact]
    public void ScatterForecastsWidenALogarithmicXAxisInSourceUnits() {
        OfficeDrawing drawing = Render(OfficeChartType.Scatter, new OfficeChartLayout(showLegend: false).WithLogarithmicValueAxes(10D, null),
            new[] { "1", "2" },
            new OfficeSeries("Sales", new[] { 1D, 2D }, new[] { 1D, 10D }, Ink, null, showMarkers: true)
                .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Linear, color: Trend).WithNativeDetails(90D, 0D, null)));

        // Forecasting 90 past x = 10 reaches 100; the axis must not read the forecast in log units (10^91).
        double largest = Texts(drawing).Select(text => double.TryParse(text.Text, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out double value) ? value : 0D).Max();
        Assert.InRange(largest, 100D, 1000D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HorizontalBarTrendlineMeetsTheBarCentresWhenCategoriesSitOnTickMarks(bool reversed) {
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false, reverseCategoryAxis: reversed).WithCategoryPointsOnTickMarks(true);
        OfficeSeries series = new OfficeSeries("Sales", new[] { 1D, 2D, 3D }, null, Ink, null, showMarkers: false)
            .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Linear, color: Trend));

        OfficeDrawing drawing = Render(OfficeChartType.BarClustered, layout, new[] { "A", "B", "C" }, series);

        double[] barCentres = Shapes(drawing).Where(shape => shape.Shape.FillColor == Ink)
            .Select(shape => shape.Y + shape.Shape.Height / 2D).OrderBy(y => y).ToArray();
        OfficePoint[] trend = TrendPoints(drawing).OrderBy(point => point.Y).ToArray();
        Assert.Equal(3, barCentres.Length);
        Assert.Equal(barCentres[0], trend[0].Y, 1);
        Assert.Equal(barCentres[2], trend[trend.Length - 1].Y, 1);
    }

    [Fact]
    public void BestFitLabelThatFitsInsideItsSliceStaysInside() {
        // The second slice spans 40 degrees centred on 3 o'clock; its 40 by 18 label box fits inside at radius 100.
        var snapshot = new OfficeChartSnapshot("", null, OfficeChartType.Pie,
            new OfficeChartData(new[] { "Big", "Small" }, new[] { new OfficeSeries("Sales", new[] { 8D, 1D }) }),
            global::OfficeIMO.OfficeLength.Points(480), global::OfficeIMO.OfficeLength.Points(300), null,
            new OfficeChartLayout(showDataLabels: true, showDataLabelValues: true, dataLabelFontSize: 12D),
            new OfficeChartRadialLayout(firstSliceAngleDegrees: 110));

        HashSet<int> roomy = OfficeChartDrawingRenderer.FindBestFitLabelsOutsideSlices(snapshot, snapshot.Layout,
            snapshot.Data.Series[0], 9D, 1D, 100D);
        HashSet<int> cramped = OfficeChartDrawingRenderer.FindBestFitLabelsOutsideSlices(snapshot, snapshot.Layout,
            snapshot.Data.Series[0], 9D, 1D, 60D);

        Assert.DoesNotContain(1, roomy);
        Assert.Contains(1, cramped);
        Assert.DoesNotContain(0, cramped);
    }

    [Fact]
    public void ErrorBarsOnALogarithmicValueAxisAreMeasuredInSourceUnits() {
        OfficeSeries series = new OfficeSeries("Sales", new[] { 10D, 100D }, new[] { 1D, 2D }, Ink, null, showMarkers: true)
            .WithErrorBars(new[] { Bars(OfficeChartErrorBarValueType.Fixed, 5D, OfficeChartErrorBarDirection.Both, endCaps: false) });

        OfficeDrawing drawing = Render(OfficeChartType.Scatter, new OfficeChartLayout(showLegend: false).WithLogarithmicValueAxes(null, 10D),
            new[] { "1", "2" }, series);

        // 10 +/- 5 spans log(15) - log(5) = 0.48 decades; 100 +/- 5 spans log(105) - log(95) = 0.043 decades.
        (OfficePoint From, OfficePoint To)[] lines = ErrorLines(drawing).OrderBy(line => line.From.X).ToArray();
        Assert.Equal(2, lines.Length);
        double ratio = Length(lines[0]) / Length(lines[1]);
        Assert.InRange(ratio, 10D, 12D);
    }

    [Fact]
    public void ErrorBarReachWidensALogarithmicValueAxis() {
        OfficeSeries series = new OfficeSeries("Sales", new[] { 10D, 100D }, new[] { 1D, 2D }, Ink, null, showMarkers: true)
            .WithErrorBars(new[] { Bars(OfficeChartErrorBarValueType.Fixed, 900D, OfficeChartErrorBarDirection.Plus, endCaps: true) });

        OfficeDrawing drawing = Render(OfficeChartType.Scatter, new OfficeChartLayout(showLegend: false).WithLogarithmicValueAxes(null, 10D),
            new[] { "1", "2" }, series);

        // The plus bar of 100 reaches 1000, so the axis gains the 1000 power, labelled "1k".
        Assert.Contains(Texts(drawing), text => text.Text == "1k");
    }

    [Fact]
    public void ScatterOnALogarithmicXAxisKeepsItsYErrorBars() {
        OfficeSeries series = new OfficeSeries("Sales", new[] { 1D, 2D }, new[] { 1D, 10D }, Ink, null, showMarkers: true)
            .WithErrorBars(new[] { Bars(OfficeChartErrorBarValueType.Fixed, 0.5D, OfficeChartErrorBarDirection.Both, endCaps: false) });

        OfficeDrawing drawing = Render(OfficeChartType.Scatter, new OfficeChartLayout(showLegend: false).WithLogarithmicValueAxes(10D, null),
            new[] { "1", "2" }, series);

        Assert.Equal(2, ErrorLines(drawing).Count());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlusErrorBarsPointTowardLargerValuesOnAReversedAxis(bool reversed) {
        OfficeSeries series = new OfficeSeries("Sales", new[] { 10D, 20D, 30D }, null, Ink, null, showMarkers: true)
            .WithErrorBars(new[] { Bars(OfficeChartErrorBarValueType.Fixed, 5D, OfficeChartErrorBarDirection.Plus, endCaps: true) });
        OfficeChartLayout layout = new OfficeChartLayout(showLegend: false);
        if (reversed) layout = layout.WithReversedValueAxes(false, true);

        OfficeDrawing drawing = Render(OfficeChartType.Line, layout, new[] { "A", "B", "C" }, series);

        // Each plus bar has its cap at the larger value: at the top normally, at the bottom on a reversed axis.
        (OfficePoint From, OfficePoint To)[] lines = ErrorLines(drawing).ToArray();
        (OfficePoint From, OfficePoint To)[] bars = lines.Where(line => Math.Abs(line.From.X - line.To.X) < 0.01D).ToArray();
        (OfficePoint From, OfficePoint To)[] caps = lines.Where(line => Math.Abs(line.From.Y - line.To.Y) < 0.01D).ToArray();
        Assert.Equal(3, bars.Length);
        Assert.Equal(3, caps.Length);
        foreach ((OfficePoint From, OfficePoint To) bar in bars) {
            double capY = caps.Single(cap => Math.Abs((cap.From.X + cap.To.X) / 2D - bar.From.X) < 0.5D).From.Y;
            double end = reversed ? Math.Max(bar.From.Y, bar.To.Y) : Math.Min(bar.From.Y, bar.To.Y);
            Assert.Equal(end, capY, 1);
        }
    }

    [Fact]
    public void StackedBarErrorBarsAreDrawnOverTheBarsStackedAboveThem() {
        OfficeColor upper = OfficeColor.Parse("#FF7F0E");
        OfficeSeries lower = new OfficeSeries("Lower", new[] { 10D }, null, Ink, null, showMarkers: false)
            .WithErrorBars(new[] { Bars(OfficeChartErrorBarValueType.Fixed, 2D, OfficeChartErrorBarDirection.Both, endCaps: true) });
        OfficeSeries top = new OfficeSeries("Upper", new[] { 10D }, null, upper, null, showMarkers: false);

        OfficeDrawing drawing = Render(OfficeChartType.ColumnStacked, new OfficeChartLayout(showLegend: false), new[] { "A" }, lower, top);

        List<OfficeDrawingShape> shapes = Shapes(drawing).ToList();
        int upperBar = shapes.FindIndex(shape => shape.Shape.FillColor == upper);
        int lastErrorLine = shapes.FindLastIndex(shape => shape.Shape.StrokeColor == ErrorBarInk);
        Assert.True(upperBar >= 0 && lastErrorLine >= 0);
        Assert.True(shapes.FindIndex(shape => shape.Shape.StrokeColor == ErrorBarInk) > upperBar,
            "Every error bar line must come after the bar stacked above it.");
    }

    [Fact]
    public void PointColoredLegendListsTheNamedTrendline() {
        OfficeSeries series = new OfficeSeries("Sales", new[] { 1D, 3D, 2D }, null, Ink,
                new OfficeColor?[] { OfficeColor.Parse("#FF0000"), OfficeColor.Parse("#00FF00"), OfficeColor.Parse("#0000FF") })
            .WithTrendline(new OfficeChartTrendline(OfficeChartTrendlineType.Linear, color: Trend).WithNativeDetails(0D, 0D, "Linear (Sales)"));

        OfficeDrawing drawing = Render(OfficeChartType.ColumnClustered, new OfficeChartLayout(showLegend: true).WithVaryColorsByPoint(true),
            new[] { "North", "South", "East" }, series);

        Assert.Contains(Texts(drawing), text => text.Text == "North");
        Assert.Contains(Texts(drawing), text => text.Text == "Linear (Sales)");
    }

    [Fact]
    public void BestFitLabelsOnAPieTooSmallForOutsideLabelsStayInsideAndAreReported() {
        var snapshot = new OfficeChartSnapshot("Small", null, OfficeChartType.Pie,
            new OfficeChartData(new[] { "North", "South", "East", "West" }, new[] { new OfficeSeries("Sales", new[] { 80D, 1D, 1D, 18D }) }),
            global::OfficeIMO.OfficeLength.Points(150), global::OfficeIMO.OfficeLength.Points(150), null,
            new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelCategoryNames: true));
        var diagnostics = new List<OfficeImageExportDiagnostic>();

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot, useMinimumCanvas: false, diagnostics);

        Assert.Contains(Texts(drawing), text => text.Text == "South");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == global::OfficeIMO.OfficeDiagnosticCodes.ChartApproximation);
    }

    [Fact]
    public void BubbleErrorBarsAreDrawn() {
        OfficeSeries series = OfficeSeries.CreateBubble("Values", new[] { 1D, 2D }, new[] { 3D, 4D }, new[] { 5D, 6D })
            .WithErrorBars(new[] { Bars(OfficeChartErrorBarValueType.Fixed, 1D, OfficeChartErrorBarDirection.Both, endCaps: false) });

        OfficeDrawing drawing = Render(OfficeChartType.Bubble, new OfficeChartLayout(showLegend: false), new[] { "1", "2" }, series);

        Assert.Equal(2, ErrorLines(drawing).Count());
    }

    private static readonly OfficeColor ErrorBarInk = OfficeColor.FromRgb(0x59, 0x59, 0x59);

    private static OfficeChartErrorBars Bars(OfficeChartErrorBarValueType kind, double amount, OfficeChartErrorBarDirection direction, bool endCaps) =>
        new OfficeChartErrorBars(kind, amount, direction, endCaps, alongCategoryAxis: false);

    private static IEnumerable<(OfficePoint From, OfficePoint To)> ErrorLines(OfficeDrawing drawing) =>
        Shapes(drawing).Where(shape => shape.Shape.StrokeColor == ErrorBarInk && shape.Shape.Kind == OfficeShapeKind.Line)
            .Select(shape => (new OfficePoint(shape.X + shape.Shape.Points[0].X, shape.Y + shape.Shape.Points[0].Y),
                new OfficePoint(shape.X + shape.Shape.Points[1].X, shape.Y + shape.Shape.Points[1].Y)));

    private static double Length((OfficePoint From, OfficePoint To) line) =>
        Math.Sqrt(Math.Pow(line.To.X - line.From.X, 2D) + Math.Pow(line.To.Y - line.From.Y, 2D));

    private static OfficeDrawing Render(OfficeChartType kind, OfficeChartLayout layout, string[] categories, params OfficeSeries[] series) =>
        OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null, kind, new OfficeChartData(categories, series),
            global::OfficeIMO.OfficeLength.Points(480), global::OfficeIMO.OfficeLength.Points(300), layout: layout));

    private static IEnumerable<OfficePoint> TrendPoints(OfficeDrawing drawing) =>
        Shapes(drawing).Where(shape => shape.Shape.StrokeColor == Trend).SelectMany(shape => shape.Shape.PathCommands
            .Where(command => command.Kind is OfficePathCommandKind.MoveTo or OfficePathCommandKind.LineTo)
            .Select(command => new OfficePoint(shape.X + command.Point.X, shape.Y + command.Point.Y)));

    private static IEnumerable<OfficeDrawingShape> Shapes(OfficeDrawing drawing) => drawing.Shapes.Concat(
        drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => Shapes(group.Drawing)));

    private static IEnumerable<OfficeDrawingText> Texts(OfficeDrawing drawing) => drawing.Elements.OfType<OfficeDrawingText>().Concat(
        drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => Texts(group.Drawing)));

    private sealed class CountingList : IReadOnlyList<double> {
        private readonly double[] _values;
        internal CountingList(double[] values) => _values = values;
        internal long Reads { get; private set; }
        public int Count => _values.Length;
        public double this[int index] { get { Reads++; return _values[index]; } }
        public IEnumerator<double> GetEnumerator() {
            foreach (double value in _values) {
                Reads++;
                yield return value;
            }
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
