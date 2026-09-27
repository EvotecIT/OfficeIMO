using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartReaderProjectionTests {
    [Fact]
    public void Snapshot_UsesLaterCategoryCacheWithoutTruncatingSeries() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Short", new[] { 1d, 2d, 3d }), new OfficeChartSeries("Long", new[] { 4d, 5d, 6d }) }));
        var series = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.LineChartSeries>().ToArray();
        series[0].GetFirstChild<C.CategoryAxisData>()!.Remove();
        var cache = series[0].GetFirstChild<C.Values>()!.Descendants<C.NumberingCache>().Single();
        cache.PointCount!.Val = 1;
        foreach (var point in cache.Elements<C.NumericPoint>().Skip(1).ToArray()) point.Remove();
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { "A", "B", "C" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 1d, 0d, 0d }, snapshot.Data.Series[0].Values);
        Assert.Equal(new[] { 4d, 5d, 6d }, snapshot.Data.Series[1].Values);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_RejectsExplicitMarkerFillWithUnresolvedConnectingLine(bool markersVisible) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Line", new[] { 1d, 2d }, null, OfficeColor.Parse("#224466"), null, markersVisible) }));
        var series = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.LineChartSeries>().Single();
        series.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!.GetFirstChild<A.SolidFill>()!.Remove();
        series.GetFirstChild<C.Marker>()!.ChartShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val = "00FF00";
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void MixedSnapshot_ValidatesWholePlotBudgetBeforeReadingIndividualLayers() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Columns", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered),
            new OfficeChartSeries("Line", new[] { 3d, 4d }, null, null, null, true, renderKind: OfficeChartKind.Line) }));
        foreach (var cache in document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.NumberingCache>()) cache.PointCount!.Val = 60000;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }
}
