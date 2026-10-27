using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartUncachedSeriesUpdateTests {
    [Theory]
    [InlineData(OfficeChartType.ColumnClustered)]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Scatter)]
    public void UpdateDataReplacesUncachedFormulaSeriesNameWithoutLosingStaticQualification(OfficeChartType kind) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        double[]? x = kind == OfficeChartType.Scatter ? new[] { 1d, 2d } : null;
        PowerPointChart chart = presentation.AddSlide().AddChart(kind,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Original", new[] { 1d, 2d }, x) }));
        SeriesText text = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!
            .Descendants<SeriesText>().Single();
        text.GetFirstChild<StringReference>()!.StringCache!.Remove();
        Assert.False(chart.TryGetOfficeSnapshot(out _));

        chart.UpdateData(new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeSeries("Updated", new[] { 3d, 4d }, x) }));

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Updated", snapshot.Data.Series.Single().Name);
        Assert.Empty(presentation.ValidateDocument());
    }
}
