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
}
