using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartLayoutQualificationTests {
    [Fact]
    public void Snapshot_RejectsRadarAxisUnitsThatTheRendererCannotDraw() {
        using var document = WordDocument.Create();
        WordChart chart = Create(document, OfficeChartKind.Radar);
        C.ValueAxis axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.AddChild(new C.MajorUnit { Val = 1 }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsSuppressedLabelsBeyondExplicitValueMaximum() {
        using var document = WordDocument.Create();
        WordChart chart = Create(document, OfficeChartKind.ColumnClustered);
        C.Chart native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        C.ValueAxis axis = native.PlotArea!.Elements<C.ValueAxis>().Single();
        axis.Scaling!.AddChild(new C.MinAxisValue { Val = 0 }, true);
        axis.Scaling!.AddChild(new C.MaxAxisValue { Val = 1.5 }, true);
        native.PlotArea.GetFirstChild<C.BarChart>()!
            .AddChild(new C.DataLabels(new C.ShowValue { Val = true }), true);
        native.GetFirstChild<C.ShowDataLabelsOverMaximum>()!.Val = true;
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        native.GetFirstChild<C.ShowDataLabelsOverMaximum>()!.Val = false;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsMajorUnitBeyondRendererTickBudget() {
        using var document = WordDocument.Create();
        WordChart chart = Create(document, OfficeChartKind.Line);
        C.ValueAxis axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.Scaling!.AddChild(new C.MinAxisValue { Val = 0 }, true);
        axis.Scaling.AddChild(new C.MaxAxisValue { Val = 100 }, true);
        axis.AddChild(new C.MajorUnit { Val = 25 }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        axis.GetFirstChild<C.MajorUnit>()!.Val = 2;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(OfficeChartKind.BarClustered)]
    [InlineData(OfficeChartKind.BarStacked)]
    [InlineData(OfficeChartKind.BarStacked100)]
    public void Snapshot_MapsHorizontalBarNumericScaleToHorizontalAxis(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var chart = Create(document, kind);
        var axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.Scaling!.Append(new C.MinAxisValue { Val = 0 }, new C.MaxAxisValue { Val = 10 });
        axis.AddChild(new C.MajorUnit { Val = 2 }, true);
        axis.GetFirstChild<C.NumberingFormat>()!.FormatCode = "0.00";
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(0d, snapshot.Layout.HorizontalAxisMinimum);
        Assert.Equal(10d, snapshot.Layout.HorizontalAxisMaximum);
        Assert.Equal(2d, snapshot.Layout.HorizontalAxisMajorUnit);
        Assert.Equal("0.00", snapshot.Layout.HorizontalAxisNumberFormat);
        Assert.Null(snapshot.Layout.VerticalAxisMaximum);
    }

    [Theory]
    [InlineData(OfficeChartKind.Line, 0)]
    [InlineData(OfficeChartKind.Scatter, 0)]
    [InlineData(OfficeChartKind.Scatter, 1)]
    public void Snapshot_RejectsUnrepresentedReversedNumericAxis(OfficeChartKind kind, int index) {
        using var document = WordDocument.Create();
        var chart = Create(document, kind);
        var axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().ElementAt(index);
        axis.Scaling!.GetFirstChild<C.Orientation>()!.Val = C.OrientationValues.MaxMin;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData("plot")]
    [InlineData("title")]
    [InlineData("legend")]
    [InlineData("topRight")]
    public void Snapshot_RejectsUnrepresentedManualLayoutAndLegendPlacement(string target) {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.Line);
        var native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        if (target == "topRight") {
            native.GetFirstChild<C.Legend>()!.LegendPosition!.Val = C.LegendPositionValues.TopRight;
        } else {
            OpenXmlCompositeElement owner = target == "plot" ? native.PlotArea! : target == "legend" ? native.GetFirstChild<C.Legend>()! : native.GetFirstChild<C.Title>()!;
            owner.RemoveAllChildren<C.Layout>();
            owner.AddChild(new C.Layout(new C.ManualLayout(new C.Left { Val = 0.2 }, new C.Top { Val = 0.2 })), true);
        }
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    private static WordChart Create(WordDocument document, OfficeChartKind kind) =>
        document.AddChart(kind, new OfficeChartData(new[] { "1", "2" }, new[] {
            new OfficeChartSeries("Values", new[] { 3d, 4d }, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null)
        }), title: "Chart");
}
