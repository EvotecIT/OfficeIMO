using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
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

    private static string Chart(string group, string before = "<c:autoTitleDeleted val=\"1\"/>", string? axes = null,
        string legend = "<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"0\"/></c:legend>", string after = "<c:dispBlanksAs val=\"gap\"/>",
        string? language = null) =>
        "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        (language == null ? string.Empty : "<c:lang val=\"" + language + "\"/>") +
        "<c:chart>" + before + "<c:plotArea><c:layout/>" + group + (axes ?? Axes()) + "</c:plotArea>" + legend +
        "<c:plotVisOnly val=\"1\"/>" + after + "</c:chart></c:chartSpace>";

    private static string Axes(string categoryPosition = "b", string valuePosition = "l",
        string categoryCrossing = "<c:crosses val=\"autoZero\"/>", string valueCrossing = "<c:crosses val=\"autoZero\"/>",
        string valueExtra = "") =>
        "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>" +
        "<c:axPos val=\"" + categoryPosition + "\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/>" +
        "<c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/>" + categoryCrossing + "<c:auto val=\"1\"/>" +
        "<c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/><c:noMultiLvlLbl val=\"0\"/></c:catAx>" +
        "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/>" +
        "<c:axPos val=\"" + valuePosition + "\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>" +
        "<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/>" +
        valueCrossing + "<c:crossBetween val=\"between\"/>" + valueExtra + "</c:valAx>";

    private static string LineGroup(string series) =>
        "<c:lineChart><c:grouping val=\"standard\"/><c:varyColors val=\"0\"/>" + series +
        "<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>";

    private static string BarGroup(string series, string direction = "col") =>
        "<c:barChart><c:barDir val=\"" + direction + "\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>" + series +
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
