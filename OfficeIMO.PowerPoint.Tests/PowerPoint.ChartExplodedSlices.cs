using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartExplodedSlicesTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void ExplodedPoint_RoundTripsAndSurvivesValueUpdate(OfficeChartKind kind) {
        var authored = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 25, 0 })
        });
        using var presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddChart(kind, authored);
        using var bytes = new MemoryStream(presentation.ToBytes());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        PowerPointChart chart = reopened.Slides.Single().Charts.Single();
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(new[] { 25, 0 }, snapshot.Data.Series.Single().PointExplosions);
        chart.UpdateData(new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 8d, 2d })
        }));
        Assert.Equal((uint)25, reopened.Slides.Single().SlidePart.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single()
            .Elements<C.DataPoint>().Single(point => point.Index!.Val!.Value == 0)
            .GetFirstChild<C.Explosion>()!.Val!.Value);
        Assert.Empty(reopened.ValidateDocument());
    }
}
