using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartProjectionReviewClosureTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Snapshot_RejectsUnsupportedSeriesAndPointColourTransforms(bool bubble, bool point) {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            bubble ? OfficeChartSeries.CreateBubble("Values", new[] { 1D, 2D }, new[] { 3D, 4D }, new[] { 5D, 6D }) :
                new OfficeChartSeries("Values", new[] { 3D, 4D })
        });
        var chart = document.AddChart(bubble ? OfficeChartKind.Bubble : OfficeChartKind.ColumnClustered, data);
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        var series = chart.ChartPart!.ChartSpace!.Descendants<OpenXmlCompositeElement>()
            .First(item => item.LocalName == "ser");
        var color = new A.RgbColorModelHex { Val = "123456" };
        color.Append(new A.HueOffset { Val = 60000 });
        var shape = new C.ChartShapeProperties(new A.SolidFill(color));
        if (point) series.AddChild(new C.DataPoint(new C.Index { Val = 0U }, shape), true);
        else {
            series.GetFirstChild<C.ChartShapeProperties>()?.Remove();
            series.AddChild(shape, true);
        }
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Snapshot_RejectsNondefaultScatterAxisSides(bool xAxis) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Scatter,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Values", new[] { 3D, 4D }, new[] { 1D, 2D })
            }));
        var axes = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().ToArray();
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        axes[xAxis ? 0 : 1].AxisPosition!.Val = xAxis ? C.AxisPositionValues.Top : C.AxisPositionValues.Right;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_PreservesIndependentSecondaryTickSettings() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Primary", new[] { 3D, 4D }),
                new OfficeChartSeries("Secondary", new[] { 5D, 6D }, null, null, null, true,
                    renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            }));
        var secondary = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        secondary.GetFirstChild<C.MajorTickMark>()!.Val = C.TickMarkValues.Inside;
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot majorSnapshot));
        Assert.Equal(OfficeChartAxisTickMark.Inside, majorSnapshot.Layout.SecondaryValueAxis!.MajorTickMark);
        secondary.GetFirstChild<C.MajorTickMark>()!.Val = C.TickMarkValues.Outside;
        secondary.GetFirstChild<C.MinorTickMark>()!.Val = C.TickMarkValues.Inside;
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot minorSnapshot));
        Assert.Equal(OfficeChartAxisTickMark.Outside, minorSnapshot.Layout.SecondaryValueAxis!.MajorTickMark);
        Assert.Equal(OfficeChartAxisTickMark.Inside, minorSnapshot.Layout.SecondaryValueAxis.MinorTickMark);
    }

    [Fact]
    public void Snapshot_RejectsOverlaidAxisTitle() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3D, 4D }) }));
        chart.SetXAxisTitle("Categories");
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        chart.ChartPart!.ChartSpace!.Descendants<C.CategoryAxis>().Single().GetFirstChild<C.Title>()!
            .AddChild(new C.Overlay { Val = true }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData("chart")]
    [InlineData("plot")]
    [InlineData("axis")]
    [InlineData("grid")]
    public void Snapshot_RejectsUnrepresentedSurfaceLineJoins(string surface) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3D, 4D }) }));
        var native = chart.ChartPart!.ChartSpace!;
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        OpenXmlCompositeElement owner = surface switch {
            "chart" => native,
            "plot" => native.GetFirstChild<C.Chart>()!.PlotArea!,
            "axis" => native.Descendants<C.ValueAxis>().Single(),
            _ => native.Descendants<C.ValueAxis>().Single().GetFirstChild<C.MajorGridlines>()!
        };
        var outline = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "123456" }), new A.Round());
        if (surface is "chart" or "plot") owner.AddChild(new C.ShapeProperties(outline), true);
        else owner.AddChild(new C.ChartShapeProperties(outline), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }
}
