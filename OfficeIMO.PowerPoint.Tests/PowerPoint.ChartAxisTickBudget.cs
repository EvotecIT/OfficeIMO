using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartAxisTickBudgetTests {
    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered, false)]
    [InlineData(OfficeChartKind.Scatter, true)]
    public void Snapshot_RejectsMajorUnitsThatCannotKeepNativeTicks(OfficeChartKind kind,
        bool horizontal) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        var data = new OfficeChartData(new[] { "0", "100" }, new[] {
            new OfficeChartSeries("Values", new[] { 0d, 100d },
                horizontal ? new[] { 0d, 100d } : null)
        });
        PowerPointChart chart = presentation.AddSlide().AddChart(kind, data);
        C.ValueAxis axis = presentation.Slides.Single().SlidePart.ChartParts.Single()
            .ChartSpace!.Descendants<C.ValueAxis>().Single(item =>
                horizontal ? item.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom :
                    item.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);
        axis.Scaling!.AddChild(new C.MinAxisValue { Val = 0 }, true);
        axis.Scaling.AddChild(new C.MaxAxisValue { Val = 100 }, true);
        axis.AddChild(new C.MajorUnit { Val = 2 }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        axis.GetFirstChild<C.MajorUnit>()!.Val = 25;
        Assert.True(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_TreatsEmptyNativeAxisDeleteAsHidden() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Values", new[] { 1d, 2d })
            }));
        C.ValueAxis axis = presentation.Slides.Single().SlidePart.ChartParts.Single()
            .ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.AddChild(new C.Delete(), true);

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.False(snapshot.Layout.ShowValueAxis);
    }
}
