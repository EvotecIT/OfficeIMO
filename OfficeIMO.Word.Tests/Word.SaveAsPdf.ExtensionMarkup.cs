using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

/// <summary>
/// Chart content repeated inside an extension list (<c>extLst</c>) is producer metadata, not chart content: the PDF
/// export reads the chart as if the extension weren't there.
/// </summary>
public sealed class WordPdfChartExtensionMarkupTests {
    private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    [Fact]
    public void ErrorBarsInsideAnExtensionDoNotStopAChartWithADataTable() {
        using var document = WordDocument.Create();
        WordChart chart = AddChart(document);
        C.PlotArea plotArea = chart.ChartPart!.ChartSpace!.Descendants<C.PlotArea>().Single();
        // A data table keeps the chart on the PDF projection, which can't draw error bars and refuses them.
        plotArea.InsertAfter(new C.DataTable(new C.ShowHorizontalBorder { Val = true }), plotArea.Elements<C.ValueAxis>().Last());
        C.BarChartSeries series = plotArea.Descendants<C.BarChartSeries>().First();
        series.AppendChild(Extension("<c:errBars><c:errDir val=\"y\"/><c:errBarType val=\"both\"/><c:errValType val=\"fixedVal\"/>" +
            "<c:noEndCap val=\"0\"/><c:val val=\"1\"/></c:errBars><c:trendline><c:trendlineType val=\"linear\"/></c:trendline>"));

        Assert.True(WordPdfConverterExtensions.TryCreateNativeWordChartSnapshot(chart, out OfficeChartSnapshot? snapshot, out string? warning), warning);
        Assert.Equal(new[] { 3d, 4d, 5d }, snapshot!.Data.Series[0].Values);
    }

    [Fact]
    public void CachedPointsInsideAnExtensionAreNotSeriesValues() {
        using var document = WordDocument.Create();
        WordChart chart = AddChart(document);
        C.BarChartSeries series = chart.ChartPart!.ChartSpace!.Descendants<C.BarChartSeries>().First();
        C.NumberingCache cache = series.GetFirstChild<C.Values>()!.Descendants<C.NumberingCache>().Single();
        cache.AppendChild(Extension("<c:numCache><c:ptCount val=\"6\"/><c:pt idx=\"5\"><c:v>99</c:v></c:pt></c:numCache>"));

        Assert.True(WordPdfConverterExtensions.TryCreateNativeWordChartSnapshot(chart, out OfficeChartSnapshot? snapshot, out string? warning), warning);
        Assert.Equal(new[] { 3d, 4d, 5d }, snapshot!.Data.Series[0].Values);
    }

    private static WordChart AddChart(WordDocument document) =>
        document.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeSeries("Values", new[] { 3d, 4d, 5d }) }));

    // An extension from a producer this reader doesn't know, holding chart elements of its own.
    private static C.ExtensionList Extension(string chartMarkup) {
        var payload = new OpenXmlUnknownElement("x", "payload", "urn:officeimo:tests:extension");
        payload.InnerXml = $"<c:root xmlns:c=\"{ChartNamespace}\">{chartMarkup}</c:root>";
        var extension = new C.Extension { Uri = "{8A5F7C33-6E6B-4C3E-9D7A-0F3E4B2C1D10}" };
        extension.AppendChild(payload);
        return new C.ExtensionList(extension);
    }
}
