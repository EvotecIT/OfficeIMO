using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartRadialLayoutTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void RadialGeometry_PersistsThroughReopenAndDataUpdate(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 7d, 3d }) });
        using PowerPointPresentation authored = PowerPointPresentation.Create();
        authored.AddSlide().AddChart(kind, data).SetRadialLayout(new OfficeChartRadialLayout(135, 75));
        using var bytes = new MemoryStream(authored.ToBytes());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        PowerPointChart chart = reopened.Slides.Single().Charts.Single();
        chart.UpdateData(data);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(135, snapshot.RadialLayout.FirstSliceAngleDegrees);
        Assert.Equal(kind == OfficeChartKind.Doughnut ? 75 : 50, snapshot.RadialLayout.DoughnutHolePercent);
        Assert.Empty(reopened.ValidateDocument());
    }
}
