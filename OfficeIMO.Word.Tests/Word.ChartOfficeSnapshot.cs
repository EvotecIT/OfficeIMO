using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartOfficeSnapshotTests {
    public static IEnumerable<object[]> SupportedKinds() => Enum.GetValues(typeof(OfficeChartKind)).Cast<OfficeChartKind>().Select(kind => new object[] { kind });

    [Theory]
    [MemberData(nameof(SupportedKinds))]
    public void OfficeSnapshot_ReadsAllSharedFamiliesAcrossNativeSaveReopen(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var expected = kind == OfficeChartKind.Bubble
            ? OfficeChartSeries.CreateBubble("Measurements", new[] { 1d, 2d }, new[] { 3d, 4d }, new[] { 5d, 6d }, OfficeColor.Parse("#224466"),
                markerOutlineColor: OfficeColor.Parse("#FF0000"), markerOutlineWidth: 2)
            : new OfficeChartSeries("Measurements", new[] { 3d, 4d }, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null, OfficeColor.Parse("#224466"));
        document.AddChart(kind, new OfficeChartData(new[] { "1", "2" }, new[] { expected }), "Measurements");
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = WordDocument.Load(bytes);
        Assert.True(reopened.Charts.Single().TryGetOfficeSnapshot(out var snapshot), kind.ToString());
        Assert.Equal(kind, snapshot.ChartKind);
        var actual = snapshot.Data.Series.Single();
        Assert.Equal(expected.Name, actual.Name);
        Assert.Equal(expected.Values, actual.Values);
        Assert.Equal(expected.XValues, actual.XValues);
        Assert.Equal(expected.BubbleSizes, actual.BubbleSizes);
        Assert.Equal(expected.Color, actual.Color);
        if (kind == OfficeChartKind.Bubble) {
            Assert.Equal(expected.MarkerOutlineColor, actual.MarkerOutlineColor);
            Assert.Equal(expected.MarkerOutlineWidth, actual.MarkerOutlineWidth);
        }
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void OfficeSnapshot_PreservesCombinationFamiliesAndSecondaryAxisAssignment() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Volume", new[] { 100d, 200d }),
            new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) }));
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = WordDocument.Load(bytes);
        Assert.True(reopened.Charts.Single().TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { OfficeChartKind.ColumnClustered, OfficeChartKind.Line }, snapshot.Data.Series.Select(series => series.RenderKind!.Value));
        Assert.Equal(new[] { OfficeChartAxisGroup.Primary, OfficeChartAxisGroup.Secondary }, snapshot.Data.Series.Select(series => series.AxisGroup));
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void OfficeSnapshot_UsesCategoryLegendIndexesForRadialCharts() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Doughnut, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 1d, 2d }) }));
        chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.GetFirstChild<C.Legend>()!.AddChild(
            new C.LegendEntry(new C.Index { Val = 0 }, new C.Delete { Val = true }), true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.True(snapshot.Data.Series.Single().ShowInLegend);
        Assert.Equal(new[] { 0 }, snapshot.Layout!.HiddenCategoryLegendIndexes);
        var drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.DoesNotContain(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "A");
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "B");
    }
}
