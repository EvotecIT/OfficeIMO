using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public class PowerPointSharedChartSeriesQualificationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeUpdate_PreservesUnprojectablePointAppearance(bool marker) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 1d }) }));
        var series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.BarChartSeries>().Single();
        var point = new C.DataPoint(new C.Index { Val = 0 });
        if (marker) point.AddChild(new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Diamond }), true);
        else point.AddChild(new C.ChartShapeProperties(new A.GradientFill(new A.GradientStopList(
            new A.GradientStop(new A.RgbColorModelHex { Val = "FF0000" }) { Position = 0 },
            new A.GradientStop(new A.RgbColorModelHex { Val = "0000FF" }) { Position = 100000 }))), true);
        series.AddChild(point, true);
        string appearance = point.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        chart.UpdateData(new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 2d }) }));
        series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.BarChartSeries>().Single();
        Assert.Equal(appearance, series.GetFirstChild<C.DataPoint>()!.OuterXml);
        Assert.Equal("2", series.GetFirstChild<C.Values>()!.Descendants<C.NumericValue>().Single().Text);
    }

    [Fact]
    public void Snapshot_IgnoresDormantMarkerOnlyLineDash() {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.Scatter, new OfficeChartData(new[] { "1" }, new[] {
            new OfficeChartSeries("Values", new[] { 2d }, new[] { 1d }, null, null, true, connectLine: false) }));
        var series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ScatterChartSeries>().Single();
        var properties = series.GetFirstChild<C.ChartShapeProperties>()!;
        var outline = properties.GetFirstChild<A.Outline>()!;
        outline.AddChild(new A.PresetDash { Val = A.PresetLineDashValues.SystemDash }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.False(snapshot.Data.Series.Single().ConnectLine);
    }
    [Theory]
    [InlineData(OfficeChartKind.ColumnStacked)]
    [InlineData(OfficeChartKind.Scatter)]
    [InlineData(OfficeChartKind.Line)]
    public void Snapshot_PreservesPlottingOrderAcrossNativeLayers(OfficeChartKind kind) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        bool mixed = kind == OfficeChartKind.Line;
        var chart = presentation.AddSlide().AddChart(kind, new OfficeChartData(new[] { "1" }, new[] {
            new OfficeChartSeries("First", new[] { 10d }, kind == OfficeChartKind.Scatter ? new[] { 1d } : null,
                color: null, pointColors: null, showMarkers: true, renderKind: mixed ? OfficeChartKind.ColumnClustered : kind),
            new OfficeChartSeries("Second", new[] { 20d }, kind == OfficeChartKind.Scatter ? new[] { 2d } : null,
                color: null, pointColors: null, showMarkers: true, renderKind: kind) }));
        var native = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!;
        var orders = native.Descendants<C.Order>().ToArray();
        orders[0].Val = 1; orders[1].Val = 0;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { "Second", "First" }, snapshot.Data.Series.Select(s => s.Name));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_RejectsSeriesAndPointGradients(bool point) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 1d }) }));
        var native = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!;
        var series = native.Descendants<C.PieChartSeries>().Single();
        var properties = new C.ChartShapeProperties(new A.GradientFill());
        if (point) series.InsertBefore(new C.DataPoint(new C.Index { Val = 0 }, properties), series.GetFirstChild<C.Values>());
        else { series.GetFirstChild<C.ChartShapeProperties>()?.Remove(); series.AddChild(properties, true); }
        string before = native.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(before, native.OuterXml);
    }
}
