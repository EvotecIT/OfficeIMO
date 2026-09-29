using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartAppearanceIntegrityTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void SharedAppearance_ExcelHidesAllRadialLegendEntries(OfficeChartKind kind) {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddWorksheet("Results").AddChart(kind, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Status", new[] { 1d, 2d, 3d }, null, null, null, true, showInLegend: false)
        }), 1, 1);
        var part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        Assert.Equal(new uint[] { 0, 1, 2 }, part.ChartSpace.Descendants<C.LegendEntry>().Select(e => e.Index!.Val!.Value));
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedAppearance_ExcelClampsAuthoredMarkerSizeAndPreflightsNativeBounds() {
        using ExcelDocument document = ExcelDocument.Create();
        var chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 1d, 2d }, null, null, null, true, markerSize: 1) }), 1, 1);
        var part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        Assert.Equal((byte)2, part.ChartSpace.Descendants<C.Size>().Single().Val!.Value);
        string before = part.ChartSpace.OuterXml;
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesLineColor(0, "FF0000", 2000));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesMarker(0, OfficeChartMarkerShape.Circle, size: 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesMarker(0, OfficeChartMarkerShape.Circle, lineWidthPoints: 2000));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetDataPointLineColor(0, 0, "FF0000", 2000));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetValueAxisGridlines(showMajor: false, showMinor: true, lineWidthPoints: 2000));
        Assert.Throws<ArgumentOutOfRangeException>(() => chart.SetSeriesTrendline(0, OfficeChartTrendlineType.Linear, lineWidthPoints: 2000));
        Assert.Equal(before, part.ChartSpace.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }
}
