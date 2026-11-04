using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartAppearanceBoundsTests {
    [Theory]
    [InlineData(2000)]
    [InlineData(0)]
    public void SharedAppearance_RejectsInvalidSeriesAndMarkerWidths(double width) {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeSeries("A", new[] { 1d }, null, null, null, true, strokeWidth: global::OfficeIMO.OfficeLength.Points(width)));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeSeries("A", new[] { 1d }, null, null, null, true, markerOutlineWidth: global::OfficeIMO.OfficeLength.Points(width)));
    }

    [Fact]
    public void SharedAppearance_AcceptsDrawingMlMaximumWidth() {
        var series = new OfficeSeries("A", new[] { 1d }, null, null, null, true, strokeWidth: global::OfficeIMO.OfficeLength.Inches(22), markerOutlineWidth: global::OfficeIMO.OfficeLength.Points(1584));
        Assert.Equal(1584, series.StrokeWidth?.ToPoints());
        Assert.Equal(1584, series.MarkerOutlineWidth?.ToPoints());
        Assert.Equal(20116800L, series.StrokeWidth!.Value.ToEmus());
        Assert.Equal(20116800L, new OfficeChartPointStyle(outlineWidth:
            global::OfficeIMO.OfficeLength.Inches(22)).OutlineWidth!.Value.ToEmus());
        var aboveMaximum = global::OfficeIMO.OfficeLength.Emus(20116801L);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeSeries("A", new[] { 1d }, null, null, null, true, strokeWidth: aboveMaximum));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeSeries("A", new[] { 1d }, null, null, null, true, markerOutlineWidth: aboveMaximum));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartPointStyle(outlineWidth: aboveMaximum));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartPointStyle(outlineWidth: global::OfficeIMO.OfficeLength.Zero));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeSeries("A", new[] { 1d }, null, null, null, true, markerOutlineWidth: global::OfficeIMO.OfficeLength.Emus(-1)));

    }
}
