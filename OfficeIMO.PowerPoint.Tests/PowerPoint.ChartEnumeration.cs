using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartEnumerationTests {
    [Fact]
    public void AllChartsIncludesChartsInsideGroupsOfALoadedSlide() {
        using var bytes = new MemoryStream();
        using (var deck = PowerPointDocument.Create(bytes, new PowerPointCreateOptions())) {
            PowerPointSlide slide = deck.AddSlide();
            PowerPointChart loose = AddChart(slide, "Loose");
            PowerPointChart grouped = AddChart(slide, "Grouped");
            slide.GroupShapes(new PowerPointShape[] { grouped, slide.AddTextBox("Caption") }, "Panel");
            deck.Save();
        }

        bytes.Position = 0;
        using var loaded = PowerPointDocument.Load(bytes, new PowerPointLoadOptions());
        PowerPointSlide reloaded = loaded.Slides[0];
        Assert.Equal(new[] { "Loose" }, reloaded.Charts.Select(chart => chart.Name));
        Assert.Equal(new[] { "Loose", "Grouped" }, reloaded.AllCharts.Select(chart => chart.Name));
        Assert.Equal(OfficeImageExportFormat.Png, reloaded.AllCharts[1].ToImage().Format);
    }

    private static PowerPointChart AddChart(PowerPointSlide slide, string name) {
        PowerPointChart chart = slide.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeSeries("Value", new[] { 1D, 2D }) }));
        chart.Name = name;
        return chart;
    }
}
