using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartCacheBoundsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_DiscoversLaterCategoriesAndRetainsLongerSeries(bool laterHasLabels) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Short", new[] { 1d, 2d, 3d }), new OfficeChartSeries("Long", new[] { 4d, 5d, 6d }) }));
        var series = chart.ChartPart!.ChartSpace!.Descendants<C.LineChartSeries>().ToArray();
        series[0].GetFirstChild<C.CategoryAxisData>()!.Remove();
        var cache = series[0].GetFirstChild<C.Values>()!.Descendants<C.NumberingCache>().Single();
        cache.PointCount!.Val = 1;
        foreach (var point in cache.Elements<C.NumericPoint>().Skip(1).ToArray()) point.Remove();
        if (!laterHasLabels) series[1].GetFirstChild<C.CategoryAxisData>()!.Remove();
        Assert.True(chart.TryGetSnapshot(out var snapshot));
        Assert.Equal(laterHasLabels ? new[] { "A", "B", "C" } : new[] { "Category 1", "Category 2", "Category 3" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 1d, 0d, 0d }, snapshot.Data.Series[0].Values);
        Assert.Equal(new[] { 4d, 5d, 6d }, snapshot.Data.Series[1].Values);
    }

    [Theory]
    [InlineData(false, 10001u)]
    [InlineData(false, uint.MaxValue)]
    [InlineData(true, 10000u)]
    [InlineData(true, uint.MaxValue)]
    public void Snapshot_RejectsOversizedCacheInsteadOfTruncatingOrReindexing(bool changeIndex, uint value) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Status", new[] { 1d, 2d }) }));
        C.NumberingCache cache = chart.ChartPart!.ChartSpace!.Descendants<C.NumberingCache>().Single();
        if (changeIndex) cache.Elements<C.NumericPoint>().Last().Index = value;
        else cache.PointCount!.Val = value;
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void Snapshot_PreservesSparseCachePositionsWithinTheSupportedBound() {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B", "C" },
            new[] { new OfficeChartSeries("Status", new[] { 1d, 2d, 3d }) }));
        C.NumberingCache cache = chart.ChartPart!.ChartSpace!.Descendants<C.NumberingCache>().Single();
        cache.Elements<C.NumericPoint>().ElementAt(1).Remove();
        Assert.True(chart.TryGetSnapshot(out WordChartSnapshot snapshot));
        Assert.Equal(new[] { 1d, 0d, 3d }, snapshot.Data.Series.Single().Values);
    }
}
