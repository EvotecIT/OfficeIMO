using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartRadialLayoutTests {
    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void RadialGeometry_PersistsThroughReopenAndDataUpdate(OfficeChartType kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Status", new[] { 7d, 3d }) });
        using PowerPointDocument authored = PowerPointDocument.Create();
        authored.AddSlide().AddChart(kind, data).SetRadialLayout(new OfficeChartRadialLayout(135, 75));
        using var bytes = new MemoryStream(authored.ToBytes());
        using PowerPointDocument reopened = PowerPointDocument.Load(bytes);
        PowerPointChart chart = reopened.Slides.Single().Charts.Single();
        chart.UpdateData(data);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(135, snapshot.RadialLayout.FirstSliceAngleDegrees);
        Assert.Equal(kind == OfficeChartType.Doughnut ? 75 : 50, snapshot.RadialLayout.DoughnutHolePercent);
        Assert.Empty(reopened.ValidateDocument());
    }
}
