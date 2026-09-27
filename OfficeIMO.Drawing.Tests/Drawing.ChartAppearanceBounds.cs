using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartAppearanceBoundsTests {
    [Theory]
    [InlineData(2000)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(0)]
    public void SharedAppearance_RejectsInvalidSeriesAndMarkerWidths(double width) {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartSeries("A", new[] { 1d }, null, null, null, true, strokeWidth: width));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartSeries("A", new[] { 1d }, null, null, null, true, markerOutlineWidth: width));
    }

    [Fact]
    public void SharedAppearance_AcceptsDrawingMlMaximumWidth() {
        var series = new OfficeChartSeries("A", new[] { 1d }, null, null, null, true, strokeWidth: 1584, markerOutlineWidth: 1584);
        Assert.Equal(1584, series.StrokeWidth);
        Assert.Equal(1584, series.MarkerOutlineWidth);
    }
}
