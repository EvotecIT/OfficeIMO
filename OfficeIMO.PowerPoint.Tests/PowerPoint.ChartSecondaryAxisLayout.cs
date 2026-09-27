using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartSecondaryAxisLayoutTests {
    [Fact]
    public void SecondaryValueAxis_PersistsThroughUpdatesAndSnapshots() {
        using var bytes = new MemoryStream();
        using var presentation = PowerPointPresentation.Create(bytes, new PowerPointCreateOptions());
        var slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Volume", new[] { 100d, 150d }),
            new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered, data);
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0, maximum: 4, majorUnit: 1,
            minorUnit: 0.5, numberFormat: "0%", majorTickMark: OfficeChartAxisTickMark.Cross,
            minorTickMark: OfficeChartAxisTickMark.Outside));
        var primary = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Left);
        primary.Scaling!.AddChild(new C.MaxAxisValue { Val = 200 }, true);
        primary.AddChild(new C.NumberingFormat { FormatCode = "0.0", SourceLinked = false }, true);
        chart.UpdateData(data);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(200, snapshot.Layout.VerticalAxisMaximum);
        Assert.Equal("0.0", snapshot.Layout.VerticalAxisNumberFormat);
        Assert.Equal(4, snapshot.Layout.SecondaryValueAxis!.Maximum);
        Assert.Equal("0%", snapshot.Layout.SecondaryValueAxis.NumberFormat);
        Assert.Empty(presentation.ValidateDocument());
        using var persistedBytes = new MemoryStream(presentation.ToBytes());
        using var reopened = PowerPointPresentation.Load(persistedBytes);
        Assert.True(reopened.Slides.Single().Charts.Single().TryGetOfficeSnapshot(out var persisted));
        Assert.Equal(0.5, persisted.Layout.SecondaryValueAxis!.MinorUnit);
        Assert.Equal(OfficeChartAxisTickMark.Cross, persisted.Layout.SecondaryValueAxis.MajorTickMark);
        Assert.Equal(OfficeChartAxisTickMark.Outside, persisted.Layout.SecondaryValueAxis.MinorTickMark);
    }
}
