using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartExplodedSlicesTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void ExplodedPoint_RoundTripsAndSurvivesValueUpdate(OfficeChartKind kind) {
        var authored = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 25, 0 })
        });
        using var document = WordDocument.Create();
        document.AddChart(kind, authored);
        using var bytes = new MemoryStream();
        document.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        WordChart chart = reopened.Charts.Single();
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(new[] { 25, 0 }, snapshot.Data.Series.Single().PointExplosions);
        chart.SetData(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 8d, 2d })
        }));
        Assert.Equal((uint)25, chart.ChartPart!.ChartSpace!.Descendants<C.PieChartSeries>()
            .Single().Elements<C.DataPoint>().Single(point => point.Index!.Val!.Value == 0)
            .GetFirstChild<C.Explosion>()!.Val!.Value);
        chart.SetData(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 8d, 2d }).WithPointExplosions(new[] { 0, 30 })
        }));
        C.PieChartSeries updated = chart.ChartPart.ChartSpace!.Descendants<C.PieChartSeries>().Single();
        Assert.Null(updated.Elements<C.DataPoint>()
            .FirstOrDefault(point => point.Index!.Val!.Value == 0)?.GetFirstChild<C.Explosion>());
        Assert.Equal((uint)30, updated.Elements<C.DataPoint>()
            .Single(point => point.Index!.Val!.Value == 1).GetFirstChild<C.Explosion>()!.Val!.Value);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void NativeSeriesExplosion_IsInheritedUnlessPointOverridesIt() {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Pie, new OfficeChartData(
            new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
                    .WithPointExplosions(new[] { 25, 0 })
            }));
        C.PieChartSeries native = chart.ChartPart!.ChartSpace!.Descendants<C.PieChartSeries>().Single();
        native.AddChild(new C.Explosion { Val = 10U }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(new[] { 25, 10 }, snapshot.Data.Series.Single().PointExplosions);
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void ShrinkingValues_IgnoresPreservedExplosionForRemovedPoint(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 0, 25 })
        }));
        chart.SetData(kind, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Status", new[] { 8d })
        }));
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Null(snapshot.Data.Series.Single().PointExplosions);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void UnsupportedNativeExplosion_PreservesEditableChartButRejectsProjection() {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Pie, new OfficeChartData(
            new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 7d, 3d }) }));
        C.PieChartSeries native = chart.ChartPart!.ChartSpace!.Descendants<C.PieChartSeries>().Single();
        native.AddChild(new C.Explosion { Val = 401U }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal((uint)401, native.GetFirstChild<C.Explosion>()!.Val!.Value);
    }
}
