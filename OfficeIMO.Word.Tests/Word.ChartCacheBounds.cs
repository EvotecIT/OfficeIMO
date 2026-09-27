using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartCacheBoundsTests {
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
