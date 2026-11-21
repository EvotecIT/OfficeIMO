using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Pins how native chart XML written by Office reaches the shared chart contract.</summary>
public sealed class PowerPointChartPlotSemanticsTests {
    [Theory]
    [InlineData("<c:dispBlanksAs val=\"gap\"/>", OfficeChartBlankDisplay.Gap)]
    [InlineData("<c:dispBlanksAs val=\"span\"/>", OfficeChartBlankDisplay.Span)]
    [InlineData("<c:dispBlanksAs val=\"zero\"/>", OfficeChartBlankDisplay.Zero)]
    [InlineData("", OfficeChartBlankDisplay.Gap)]
    public void EmptyCellsStayMissingAndCarryTheChartBlankSetting(string blanks, OfficeChartBlankDisplay expected) {
        using PowerPointDocument presentation = LoadChart(Chart(LineGroup(Series(0, "Sales", 90, 95, null, 101)), after: blanks));

        OfficeChartSnapshot snapshot = Snapshot(presentation);

        Assert.Equal(new[] { 90D, 95D, double.NaN, 101D }, snapshot.Data.Series.Single().Values);
        Assert.Equal(expected, snapshot.Layout.BlankDisplay);
    }

    [Theory]
    [InlineData(1, "0", null, "Revenue")]
    [InlineData(2, "0", null, "Chart Title")]
    [InlineData(2, "0", "en-GB", "Chart Title")]
    [InlineData(2, "0", "de-DE", null)]
    [InlineData(1, "1", null, null)]
    public void TitleWithoutTextShowsWhatOfficeShows(int seriesCount, string autoTitleDeleted, string? language, string? expected) {
        string series = Series(0, "Revenue", 3, 4, 5) + (seriesCount > 1 ? Series(1, "Cost", 1, 2, 3) : string.Empty);
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(series),
            before: "<c:title><c:overlay val=\"0\"/></c:title><c:autoTitleDeleted val=\"" + autoTitleDeleted + "\"/>",
            language: language));

        Assert.Equal(expected, Snapshot(presentation).Title);
    }

    [Theory]
    [InlineData("col", "<c:crosses val=\"max\"/>", "<c:crosses val=\"autoZero\"/>",
        OfficeChartAxisCrossingPosition.Maximum, null, OfficeChartAxisCrossingPosition.AutoZero, null)]
    [InlineData("col", "<c:crosses val=\"autoZero\"/>", "<c:crosses val=\"max\"/>",
        OfficeChartAxisCrossingPosition.AutoZero, null, OfficeChartAxisCrossingPosition.Maximum, null)]
    [InlineData("col", "<c:crossesAt val=\"-10\"/>", "<c:crosses val=\"min\"/>",
        OfficeChartAxisCrossingPosition.AutoZero, -10D, OfficeChartAxisCrossingPosition.Minimum, null)]
    [InlineData("bar", "<c:crossesAt val=\"-10\"/>", "<c:crosses val=\"max\"/>",
        OfficeChartAxisCrossingPosition.Maximum, null, OfficeChartAxisCrossingPosition.AutoZero, -10D)]
    public void AxisCrossingPlacesTheAxisThatDeclaresIt(string direction, string categoryCrossing, string valueCrossing,
        OfficeChartAxisCrossingPosition horizontal, double? horizontalAt, OfficeChartAxisCrossingPosition vertical, double? verticalAt) {
        bool bars = direction == "bar";
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(Series(0, "Sales", 30, -20, 45), direction),
            axes: Axes(categoryPosition: bars ? "l" : "b", valuePosition: bars ? "b" : "l",
                categoryCrossing: categoryCrossing, valueCrossing: valueCrossing)));

        OfficeChartLayout layout = Snapshot(presentation).Layout;

        Assert.Equal((horizontal, horizontalAt, vertical, verticalAt),
            (layout.HorizontalAxisCrossingPosition, layout.HorizontalAxisCrossesAt, layout.VerticalAxisCrossingPosition, layout.VerticalAxisCrossesAt));
    }

    [Fact]
    public void ValueAxisDisplayUnitsDivideTheAxisLabels() {
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(Series(0, "Sales", 3000, 2000, 4500)),
            axes: Axes(valueExtra: "<c:dispUnits><c:builtInUnit val=\"thousands\"/><c:dispUnitsLbl/></c:dispUnits>")));

        OfficeChartLayout layout = Snapshot(presentation).Layout;

        Assert.Equal(1000D, layout.VerticalAxisDisplayUnitDivisor);
        Assert.Equal("Thousands", layout.VerticalAxisDisplayUnitLabel);
    }

    [Fact]
    public void DisplayUnitLabelLinkedToACellShowsItsCachedText() {
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(Series(0, "Sales", 3000, 2000, 4500)),
            axes: Axes(valueExtra: "<c:dispUnits><c:builtInUnit val=\"thousands\"/><c:dispUnitsLbl><c:tx><c:strRef><c:f>Sheet1!$D$1</c:f>" +
                "<c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>in thousands</c:v></c:pt></c:strCache></c:strRef></c:tx></c:dispUnitsLbl></c:dispUnits>")));

        Assert.Equal("in thousands", Snapshot(presentation).Layout.VerticalAxisDisplayUnitLabel);
    }

    [Fact]
    public void TopRightLegendKeepsItsPosition() {
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(Series(0, "A", 1, 2, 3) + Series(1, "B", 2, 3, 4)),
            legend: "<c:legend><c:legendPos val=\"tr\"/><c:overlay val=\"0\"/></c:legend>"));

        Assert.Equal(OfficeChartLegendPosition.TopRight, Snapshot(presentation).Layout.LegendPosition);
    }

    // Checked against Excel 16: with no c:title, a one-series chart is titled by its series name only
    // when c:autoTitleDeleted is present and off; a missing element or a second series gives no title.
    [Theory]
    [InlineData(1, "<c:autoTitleDeleted val=\"0\"/>", "Revenue")]
    [InlineData(1, "", null)]
    [InlineData(2, "<c:autoTitleDeleted val=\"0\"/>", null)]
    public void ChartWithoutTitleElementShowsWhatExcelShows(int seriesCount, string autoTitleDeleted, string? expected) {
        string series = Series(0, "Revenue", 3, 4, 5) + (seriesCount > 1 ? Series(1, "Cost", 1, 2, 3) : string.Empty);
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(series), before: autoTitleDeleted));

        Assert.Equal(expected, Snapshot(presentation).Title);
    }

    [Theory]
    [InlineData("col", "<c:logBase val=\"10\"/><c:orientation val=\"minMax\"/>", null, 10D, false, false)]
    [InlineData("col", "<c:orientation val=\"maxMin\"/>", null, null, false, true)]
    [InlineData("bar", "<c:logBase val=\"2\"/><c:orientation val=\"maxMin\"/>", 2D, null, true, false)]
    public void ValueAxisScalesReachThePhysicalAxis(string direction, string scaling, double? horizontalBase, double? verticalBase,
        bool reverseHorizontal, bool reverseVertical) {
        bool bars = direction == "bar";
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(Series(0, "Sales", 3, 40, 500), direction),
            axes: Axes(categoryPosition: bars ? "l" : "b", valuePosition: bars ? "b" : "l", valueScaling: scaling)));

        OfficeChartLayout layout = Snapshot(presentation).Layout;

        Assert.Equal((horizontalBase, verticalBase, reverseHorizontal, reverseVertical),
            (layout.HorizontalAxisLogBase, layout.VerticalAxisLogBase, layout.ReverseHorizontalValueAxis, layout.ReverseVerticalValueAxis));
    }

    [Theory]
    [InlineData(10D, null)]
    [InlineData(0.5D, "A stack whose total is below 1 cannot be drawn on a logarithmic value axis.")]
    public void StackedColumnsOnALogarithmicAxisDrawOrReportWhyNot(double firstValue, string? reason) {
        string group = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"stacked\"/><c:varyColors val=\"0\"/>" +
            Series(0, "A", firstValue, 10, 10) + Series(1, "B", 90, 90, 90) +
            "<c:gapWidth val=\"150\"/><c:overlap val=\"100\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>";
        using PowerPointDocument presentation = LoadChart(Chart(group,
            axes: Axes(valueScaling: "<c:logBase val=\"10\"/><c:orientation val=\"minMax\"/>")));

        OfficeImageExportResult image = presentation.Slides[0].ToImage(OfficeImageExportFormat.Png);
        PdfDocumentConversionResult pdf = presentation.ConvertToPdfWithEvidence();

        if (reason == null) {
            Assert.DoesNotContain(image.Diagnostics, diagnostic => diagnostic.Code == "unsupported-chart");
            Assert.DoesNotContain(pdf.Warnings, warning => warning.Code == "unsupported-chart");
        } else {
            Assert.Contains(reason, Assert.Single(image.Diagnostics, diagnostic => diagnostic.Code == "unsupported-chart").Message);
            Assert.Contains(reason, Assert.Single(pdf.Warnings, warning => warning.Code == "unsupported-chart").Message);
        }
    }

    [Fact]
    public void ManualLayoutsReachTheLayout() {
        const string plot = "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/>" +
            "<c:x val=\"0.2\"/><c:y val=\"0.25\"/><c:w val=\"0.5\"/><c:h val=\"0.6\"/></c:manualLayout></c:layout>";
        const string legend = "<c:legend><c:legendPos val=\"r\"/><c:layout><c:manualLayout><c:x val=\"0.05\"/><c:y val=\"-0.1\"/>" +
            "</c:manualLayout></c:layout><c:overlay val=\"0\"/></c:legend>";
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(Series(0, "A", 1, 2, 3) + Series(1, "B", 2, 3, 4)),
            plotLayout: plot, legend: legend));

        OfficeChartLayout layout = Snapshot(presentation).Layout;

        OfficeChartManualLayout placedPlot = Assert.IsType<OfficeChartManualLayout>(layout.PlotAreaLayout);
        Assert.Equal((0.2D, 0.25D, (double?)0.5D, (double?)0.6D, OfficeChartLayoutMode.Edge, OfficeChartLayoutTarget.Inner),
            (placedPlot.X, placedPlot.Y, placedPlot.Width, placedPlot.Height, placedPlot.XMode, placedPlot.Target));
        OfficeChartManualLayout placedLegend = Assert.IsType<OfficeChartManualLayout>(layout.LegendLayout);
        Assert.Equal((0.05D, -0.1D, OfficeChartLayoutMode.Factor, (double?)null),
            (placedLegend.X, placedLegend.Y, placedLegend.XMode, placedLegend.Width));
        Assert.Null(layout.TitleLayout);
    }

    [Theory]
    [InlineData(1, true)]
    [InlineData(2, false)]
    public void VaryColorsAppliesToASingleSeries(int seriesCount, bool expected) {
        string series = Series(0, "A", 1, 2, 3) + (seriesCount > 1 ? Series(1, "B", 2, 3, 4) : string.Empty);
        using PowerPointDocument presentation = LoadChart(Chart(BarGroup(series, varyColors: true)));

        OfficeChartSnapshot snapshot = Snapshot(presentation);

        Assert.Equal(expected, snapshot.Layout.VaryColorsByPoint);
        if (expected) Assert.Equal(3, Enumerable.Range(0, 3).Select(snapshot.Style.GetSeriesColor).Distinct().Count());
    }

    [Fact]
    public void LabelsBeyondAnExplicitMaximumNoLongerBlockTheSnapshot() {
        string group = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" + Series(0, "A", 5, 50, 20) +
            "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/>" +
            "<c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls><c:gapWidth val=\"150\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>";
        using PowerPointDocument presentation = LoadChart(Chart(group,
            axes: Axes(valueScaling: "<c:orientation val=\"minMax\"/><c:max val=\"30\"/><c:min val=\"0\"/>"),
            after: "<c:dispBlanksAs val=\"gap\"/><c:showDLblsOverMax val=\"0\"/>"));

        Assert.False(Snapshot(presentation).Layout.ShowDataLabelsOverMaximum);
    }

    [Fact]
    public void ChartsWrittenWithoutATitleDeleteTheAutomaticTitle() {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartType.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Revenue", new[] { 1D }) }));
        DocumentFormat.OpenXml.Drawing.Charts.Chart native = presentation.Slides.Single().SlidePart.ChartParts.Single()
            .ChartSpace!.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>()!;

        Assert.True(native.AutoTitleDeleted?.Val?.Value);
        chart.SetTitle("Revenue by quarter");
        Assert.False(native.AutoTitleDeleted?.Val?.Value);
    }

    private static OfficeChartSnapshot Snapshot(PowerPointDocument presentation) {
        Assert.True(presentation.Slides[0].Charts.Single().TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        return snapshot;
    }

    private static PowerPointDocument LoadChart(string chartXml) {
        var stream = new MemoryStream();
        using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
            presentation.AddSlide().AddChart(OfficeChartType.ColumnClustered,
                new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Value", new[] { 1D }) }));
            presentation.Save();
        }

        stream.Position = 0;
        using (PresentationDocument package = PresentationDocument.Open(stream, true)) {
            ChartPart part = package.PresentationPart!.SlideParts.Single().ChartParts.Single();
            using Stream chart = part.GetStream(FileMode.Create);
            byte[] bytes = new UTF8Encoding(false).GetBytes(chartXml);
            chart.Write(bytes, 0, bytes.Length);
        }

        stream.Position = 0;
        return PowerPointDocument.Load(stream);
    }

    private static string Chart(string group, string before = "<c:autoTitleDeleted val=\"1\"/>", string? axes = null, string plotLayout = "<c:layout/>",
        string legend = "<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"0\"/></c:legend>", string after = "<c:dispBlanksAs val=\"gap\"/>",
        string? language = null) =>
        "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        (language == null ? string.Empty : "<c:lang val=\"" + language + "\"/>") +
        "<c:chart>" + before + "<c:plotArea>" + plotLayout + group + (axes ?? Axes()) + "</c:plotArea>" + legend +
        "<c:plotVisOnly val=\"1\"/>" + after + "</c:chart></c:chartSpace>";

    private static string Axes(string categoryPosition = "b", string valuePosition = "l",
        string categoryCrossing = "<c:crosses val=\"autoZero\"/>", string valueCrossing = "<c:crosses val=\"autoZero\"/>",
        string valueExtra = "", string valueScaling = "<c:orientation val=\"minMax\"/>") =>
        "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>" +
        "<c:axPos val=\"" + categoryPosition + "\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/>" +
        "<c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/>" + categoryCrossing + "<c:auto val=\"1\"/>" +
        "<c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/><c:noMultiLvlLbl val=\"0\"/></c:catAx>" +
        "<c:valAx><c:axId val=\"2\"/><c:scaling>" + valueScaling + "</c:scaling><c:delete val=\"0\"/>" +
        "<c:axPos val=\"" + valuePosition + "\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>" +
        "<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/>" +
        valueCrossing + "<c:crossBetween val=\"between\"/>" + valueExtra + "</c:valAx>";

    private static string LineGroup(string series) =>
        "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + series +
        "<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>";

    private static string BarGroup(string series, string direction = "col", bool varyColors = false) =>
        "<c:barChart><c:barDir val=\"" + direction + "\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"" + (varyColors ? "1" : "0") + "\"/>" + series +
        "<c:gapWidth val=\"150\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>";

    // A null value is an empty cell: Office leaves its point out of the cache.
    private static string Series(int index, string name, params double?[] values) {
        var xml = new StringBuilder();
        xml.Append("<c:ser><c:idx val=\"").Append(index).Append("\"/><c:order val=\"").Append(index).Append("\"/>")
            .Append("<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>")
            .Append(name).Append("</c:v></c:pt></c:strCache></c:strRef></c:tx>")
            .Append("<c:cat><c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val=\"").Append(values.Length).Append("\"/>");
        for (int i = 0; i < values.Length; i++) xml.Append("<c:pt idx=\"").Append(i).Append("\"><c:v>C").Append(i + 1).Append("</c:v></c:pt>");
        xml.Append("</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode>")
            .Append("<c:ptCount val=\"").Append(values.Length).Append("\"/>");
        for (int i = 0; i < values.Length; i++) {
            if (values[i] is double value)
                xml.Append("<c:pt idx=\"").Append(i).Append("\"><c:v>")
                    .Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append("</c:v></c:pt>");
        }
        return xml.Append("</c:numCache></c:numRef></c:val></c:ser>").ToString();
    }
}
