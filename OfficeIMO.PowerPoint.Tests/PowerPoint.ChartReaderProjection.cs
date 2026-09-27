using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartReaderProjectionTests {
    [Fact]
    public void MixedSnapshot_RejectsAggregatePaddingExpansionAcrossLayers() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Long", new[] { 1d }), new OfficeChartSeries("Short", new[] { 2d }) }));
        var plot = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var original = plot.GetFirstChild<C.LineChart>()!;
        original.Elements<C.LineChartSeries>().First().GetFirstChild<C.Values>()!.Descendants<C.PointCount>().Single().Val = 25000;
        for (uint layerIndex = 1; layerIndex < 3; layerIndex++) {
            var copy = (C.LineChart)original.CloneNode(true);
            uint position = layerIndex * 2;
            foreach (var series in copy.Elements<C.LineChartSeries>()) { series.Index!.Val = position; series.Order!.Val = position++; }
            plot.InsertBefore(copy, original);
        }
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedSnapshot_RejectsIncompatibleCategoryCachesWithoutDroppingSeries(bool differentLength) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Columns", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered),
            new OfficeChartSeries("Line", new[] { 3d, 4d }, null, null, null, true, renderKind: OfficeChartKind.Line) }));
        var series = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.LineChartSeries>().Single();
        if (differentLength) {
            series.GetFirstChild<C.Values>()!.Descendants<C.PointCount>().Single().Val = 3;
            series.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.PointCount>().Single().Val = 3;
        } else series.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringPoint>().Last().NumericValue!.Text = "Different";
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

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
