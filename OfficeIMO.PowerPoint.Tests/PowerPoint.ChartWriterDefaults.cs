using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public class PowerPointChartWriterDefaultsTests {
    [Fact]
    public void StyledColumnSeriesHaveNoOutlineUnlessAsked() {
        using var presentation = PowerPointPresentation.Create();
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Sales", new[] { 1d, 2d }, null, OfficeColor.Parse("#156082")) }));
        C.BarChartSeries series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!
            .Descendants<C.BarChartSeries>().Single();
        Assert.Null(series.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>());
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.NotEqual(true, snapshot.Data.Series.Single().SeriesStyle?.ShowOutline);
    }

    [Fact]
    public void DoughnutSeriesFillTurnsVaryColorsOff() {
        using var presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "Direct", "Partner" }, new[] { new OfficeChartSeries("Share", new[] { 60d, 40d }) }));
        C.DoughnutChart doughnut = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!
            .Descendants<C.DoughnutChart>().Single();
        Assert.True(doughnut.GetFirstChild<C.VaryColors>()!.Val!.Value);
        chart.SetSeriesFillColor(0, "156082");
        // A series fill colours every slice in PowerPoint; c:varyColors must agree for other readers.
        Assert.False(doughnut.GetFirstChild<C.VaryColors>()!.Val!.Value);
    }
}
