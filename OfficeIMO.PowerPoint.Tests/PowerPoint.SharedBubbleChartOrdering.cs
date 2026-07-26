using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public class PowerPointSharedBubbleChartOrderingTests {
    [Fact]
    public void BubbleChart_SnapshotUsesNativeSeriesOrder() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create(new MemoryStream());
        var data = new OfficeChartData(new[] { "1" }, new[] {
            OfficeChartSeries.CreateBubble(
                "First", new[] { 1D }, new[] { 10D }, new[] { 4D }),
            OfficeChartSeries.CreateBubble(
                "Second", new[] { 2D }, new[] { 20D }, new[] { 9D })
        });
        PowerPointChart chart = presentation.AddSlide().AddChart(
            OfficeChartKind.Bubble, data);
        ChartPart chartPart = presentation.Slides[0].SlidePart
            .ChartParts.Single();
        C.BubbleChartSeries[] nativeSeries = chartPart.ChartSpace!
            .Descendants<C.BubbleChartSeries>().ToArray();
        nativeSeries[0].Order!.Val = 1U;
        nativeSeries[1].Order!.Val = 0U;

        Assert.True(chart.TryGetOfficeSnapshot(
            out OfficeChartSnapshot snapshot));
        Assert.Equal(new[] { "Second", "First" },
            snapshot.Data.Series.Select(series => series.Name));
        Assert.Equal(new[] { 20D, 10D },
            snapshot.Data.Series.Select(series => series.Values[0]));

        nativeSeries[0].Order.Val = 0U;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        nativeSeries[0].Order.Remove();
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }
}
