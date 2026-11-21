using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Moving-average trendline points.</summary>
public sealed class DrawingChartMovingAverageTests {
    [Fact]
    public void WindowsAfterALargeValueLeavesAreExact() {
        double[] averages = OfficeChartDrawingRenderer.GetMovingAverages(new[] { 18014398509481984D, 1D, 1D, 1D }, 2);

        Assert.Equal(new[] { 9007199254740992D, 1D, 1D }, averages);
    }
}
