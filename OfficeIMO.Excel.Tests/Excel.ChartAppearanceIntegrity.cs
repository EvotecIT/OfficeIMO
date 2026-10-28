using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartAppearanceIntegrityTests {
    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void SharedAppearance_ExcelHidesAllRadialLegendEntries(OfficeChartType kind) {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddSheet("Results").AddChart(kind, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeSeries("Status", new[] { 1d, 2d, 3d }, null, null, null, true, showInLegend: false)
        }), 1, 1);
        var part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        Assert.Equal(new uint[] { 0, 1, 2 }, part.ChartSpace.Descendants<C.LegendEntry>().Select(e => e.Index!.Val!.Value));
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedAppearance_ExcelClampsAuthoredMarkerSizeAndPreflightsNativeBounds() {
        using ExcelDocument document = ExcelDocument.Create();
        var chart = document.AddSheet("Results").AddChart(OfficeChartType.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Status", new[] { 1d, 2d }, null, null, null, true, markerSize: 1) }), 1, 1);
        var part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        Assert.Equal((byte)2, part.ChartSpace.Descendants<C.Size>().Single().Val!.Value);
        string before = part.ChartSpace.OuterXml;
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesLineColor(0, OfficeColor.Parse("FF0000"), OfficeLength.Points(2000)));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesMarker(0, OfficeChartMarkerShape.Circle, size: OfficeLength.Points(1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesMarker(0, OfficeChartMarkerShape.Circle, lineWidth: OfficeLength.Points(2000)));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetDataPointLineColor(0, 0, OfficeColor.Parse("FF0000"), OfficeLength.Points(2000)));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetValueAxisGridlines(showMajor: false, showMinor: true, lineWidth: OfficeLength.Points(2000)));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesTrendline(0, OfficeChartTrendlineType.Linear, lineWidth: OfficeLength.Points(2000)));
        Assert.Equal(before, part.ChartSpace.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }
}
