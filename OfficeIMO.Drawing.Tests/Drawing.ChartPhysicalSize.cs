using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class ChartPhysicalSizeTests {
    [Fact]
    public void SnapshotRetainsExactPhysicalDimensionsThroughResizeAndRendering() {
        var data = new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Sales", new[] { 4D }) });
        var snapshot = new OfficeChartSnapshot("Sales", null, OfficeChartType.ColumnClustered, data,
            OfficeLength.Emus(1270001), OfficeLength.Inches(2));
        var resized = snapshot.WithSize(OfficeLength.Centimeters(5), snapshot.Height);

        Assert.Equal(1270001, snapshot.Width.ToEmus());
        Assert.Equal(OfficeLength.Inches(2), resized.Height);
        Assert.Same(snapshot.Data, resized.Data);
        Assert.Same(snapshot.Style, resized.Style);
        Assert.Same(snapshot.Layout, resized.Layout);
        var drawing = OfficeChartDrawingRenderer.Render(resized, useMinimumCanvas: false);
        Assert.Equal(resized.Width.ToPoints(), drawing.Width);
        Assert.Equal(resized.Height.ToPoints(), drawing.Height);
        Assert.Throws<ArgumentOutOfRangeException>(() => snapshot.WithSize(OfficeLength.Zero, snapshot.Height));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartSnapshot("Invalid", null,
            OfficeChartType.ColumnClustered, data, OfficeLength.Points(-1), snapshot.Height));
        Assert.Equal(1270001, snapshot.Width.ToEmus());
    }
}
