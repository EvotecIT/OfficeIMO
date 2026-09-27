using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartRadialLayoutTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void RadialGeometry_PersistsThroughReopenAndDataUpdate(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 7d, 3d }) });
        using WordDocument authored = WordDocument.Create();
        authored.AddChart(kind, data).SetRadialLayout(new OfficeChartRadialLayout(135, 75));
        using var bytes = new MemoryStream();
        authored.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        WordChart chart = reopened.Charts.Single();
        chart.SetData(kind, data);
        Assert.True(chart.TryGetSnapshot(out WordChartSnapshot snapshot));
        Assert.Equal(135, snapshot.RadialLayout.FirstSliceAngleDegrees);
        Assert.Equal(kind == OfficeChartKind.Doughnut ? 75 : 50, snapshot.RadialLayout.DoughnutHolePercent);
        Assert.Empty(reopened.ValidateDocument());
    }
}
