using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointSharedChartTextStyleTests {
    [Fact]
    public void SharedSnapshot_PreservesExplicitNativeChartFonts() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetTitleTextStyle(fontName: "Arial")
            .SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Arial", snapshot.Style.FontFamily);
        Assert.Equal("Arial", snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotProjectTitleFontOntoChartBodyText() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetTitleTextStyle(fontName: "Georgia")
            .SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Arial", snapshot.Style.FontFamily);
        Assert.Equal("Georgia", snapshot.Style.TitleFontFamily);
    }
}
