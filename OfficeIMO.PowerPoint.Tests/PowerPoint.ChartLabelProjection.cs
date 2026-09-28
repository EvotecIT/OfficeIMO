using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartLabelProjectionTests {
    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Scatter)]
    public void SharedSnapshot_PreservesNativeDataLabelsOnNonRadialCharts(OfficeChartKind kind) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        var series = new OfficeChartSeries("Values", new[] { 3d, 4d },
            kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null);
        PowerPointChart chart = presentation.AddSlide().AddChart(kind,
            new OfficeChartData(new[] { "A", "B" }, new[] { series })).SetDataLabels(showValue: true);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.True(snapshot.Layout.ShowDataLabels);
        Assert.True(snapshot.Layout.ShowDataLabelValues);
        Assert.Contains(OfficeChartDrawingRenderer.Render(snapshot).Elements.OfType<OfficeDrawingText>(),
            text => text.Text.Contains("3"));
    }
}
