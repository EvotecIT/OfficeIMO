using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public class PowerPointAllSeverityBatch21Tests {
    [Fact]
    public void CategorySeriesXValuesCannotBypassValueCountValidation() {
        PowerPointChartSeries malformed = new(
            "Series",
            new[] { 1D },
            new[] { 1D });

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            new PowerPointChartData(
                new[] { "First", "Second" },
                new[] { malformed }));

        Assert.Equal("series", exception.ParamName);
    }

    [Fact]
    public void ScatterSnapshotAllowsSeriesWithDifferentPointCounts() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        PowerPointSlide slide = presentation.AddSlide();
        var data = new PowerPointScatterChartData(new[] {
            new PowerPointScatterChartSeries(
                "Long",
                new[] { 1D, 2D, 3D },
                new[] { 10D, 20D, 30D }),
            new PowerPointScatterChartSeries(
                "Short",
                new[] { 4D, 5D },
                new[] { 40D, 50D })
        });
        PowerPointChart chart = slide.AddScatterChart(data);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        Assert.Equal(new[] { 3, 2 }, snapshot.Data.Series.Select(series => series.Values.Count).ToArray());
        Assert.Equal(new[] { 3, 2 }, snapshot.Data.Series.Select(series => series.XValues!.Count).ToArray());
    }
}
