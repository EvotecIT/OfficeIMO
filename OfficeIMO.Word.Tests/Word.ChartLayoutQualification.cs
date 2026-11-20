using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartLayoutQualificationTests {
    [Fact]
    public void Snapshot_RejectsRadarAxisUnitsThatTheRendererCannotDraw() {
        using var document = WordDocument.Create();
        WordChart chart = Create(document, OfficeChartType.Radar);
        C.ValueAxis axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.AddChild(new C.MajorUnit { Val = 1 }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_ProjectsLabelsHiddenBeyondExplicitValueMaximum() {
        using var document = WordDocument.Create();
        WordChart chart = Create(document, OfficeChartType.ColumnClustered);
        C.Chart native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        C.ValueAxis axis = native.PlotArea!.Elements<C.ValueAxis>().Single();
        axis.Scaling!.AddChild(new C.MinAxisValue { Val = 0 }, true);
        axis.Scaling!.AddChild(new C.MaxAxisValue { Val = 1.5 }, true);
        native.PlotArea.GetFirstChild<C.BarChart>()!
            .AddChild(new C.DataLabels(new C.ShowValue { Val = true }), true);
        native.GetFirstChild<C.ShowDataLabelsOverMaximum>()!.Val = true;
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        native.GetFirstChild<C.ShowDataLabelsOverMaximum>()!.Val = false;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.False(snapshot.Layout.ShowDataLabelsOverMaximum);
    }

    [Fact]
    public void Snapshot_RejectsMajorUnitBeyondRendererTickBudget() {
        using var document = WordDocument.Create();
        WordChart chart = Create(document, OfficeChartType.Line);
        C.ValueAxis axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.Scaling!.AddChild(new C.MinAxisValue { Val = 0 }, true);
        axis.Scaling.AddChild(new C.MaxAxisValue { Val = 100 }, true);
        axis.AddChild(new C.MajorUnit { Val = 25 }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        axis.GetFirstChild<C.MajorUnit>()!.Val = 2;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsDerivedMinorTicksBeyondRendererBudget() {
        using var document = WordDocument.Create();
        WordChart chart = Create(document, OfficeChartType.Line);
        C.ValueAxis axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.Scaling!.AddChild(new C.MinAxisValue { Val = 0 }, true);
        axis.Scaling.AddChild(new C.MaxAxisValue { Val = 100 }, true);
        axis.AddChild(new C.MajorUnit { Val = 4 }, true);
        axis.AddChild(new C.MinorTickMark { Val = C.TickMarkValues.Outside }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(OfficeChartType.BarClustered)]
    [InlineData(OfficeChartType.BarStacked)]
    [InlineData(OfficeChartType.BarStacked100)]
    public void Snapshot_MapsHorizontalBarNumericScaleToHorizontalAxis(OfficeChartType kind) {
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
    [InlineData(OfficeChartType.Line, 0)]
    [InlineData(OfficeChartType.Scatter, 0)]
    [InlineData(OfficeChartType.Scatter, 1)]
    public void Snapshot_ProjectsReversedNumericAxis(OfficeChartType kind, int index) {
        using var document = WordDocument.Create();
        var chart = Create(document, kind);
        var axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().ElementAt(index);
        axis.Scaling!.GetFirstChild<C.Orientation>()!.Val = C.OrientationValues.MaxMin;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        bool horizontal = kind == OfficeChartType.Scatter && index == 0;
        Assert.True(horizontal ? snapshot.Layout.ReverseHorizontalValueAxis : snapshot.Layout.ReverseVerticalValueAxis);
    }

    private static WordChart Create(WordDocument document, OfficeChartType kind) =>
        document.AddChart(kind, new OfficeChartData(new[] { "1", "2" }, new[] {
            new OfficeSeries("Values", new[] { 3d, 4d }, kind == OfficeChartType.Scatter ? new[] { 1d, 2d } : null)
        }), new OfficeIMO.Word.WordChartOptions { Title = "Chart" });
}
