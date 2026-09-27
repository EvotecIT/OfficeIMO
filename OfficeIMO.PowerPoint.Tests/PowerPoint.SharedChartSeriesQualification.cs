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
    public void Snapshot_QualifiesZeroWidthPointOutlineVisibility(bool hidden) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.Pie, new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 2d }) }));
        var series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.PieChartSeries>().Single();
        var outline = new A.Outline { Width = 0 };
        if (hidden) outline.Append(new A.NoFill());
        else outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "123456" }));
        series.AddChild(new C.DataPoint(new C.Index { Val = 0 }, new C.ChartShapeProperties(outline)), true);
        Assert.Equal(hidden, chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsExplicitZeroWidthBubbleOutline() {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.Bubble, new OfficeChartData(new[] { "1" }, new[] {
            OfficeChartSeries.CreateBubble("Values", new[] { 1d }, new[] { 2d }, new[] { 3d }, OfficeColor.Parse("#123456")) }));
        var properties = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.BubbleChartSeries>().Single().ChartShapeProperties!;
        var outline = properties.GetFirstChild<A.Outline>()!;
        outline.RemoveAllChildren(); outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "123456" })); outline.Width = 0;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(OfficeChartKind.Line, false)]
    [InlineData(OfficeChartKind.Scatter, false)]
    [InlineData(OfficeChartKind.Line, true)]
    [InlineData(OfficeChartKind.ColumnClustered, false)]
    public void Snapshot_RejectsExplicitZeroWidthVisibleOutlines(OfficeChartKind kind, bool marker) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var chart = presentation.AddSlide().AddChart(kind, new OfficeChartData(new[] { "1" }, new[] {
            new OfficeChartSeries("Values", new[] { 2d }, kind == OfficeChartKind.Scatter ? new[] { 1d } : null, OfficeColor.Parse("#123456")) }));
        var series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants().OfType<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(item => item.LocalName == "ser");
        if (marker) series.GetFirstChild<C.Marker>()!.AddChild(new C.ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "123456" })) { Width = 0 }), true);
        else {
            var outline = series.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!;
            outline.RemoveAllChildren();
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "123456" }));
            outline.Width = 0;
        }
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false, OfficeChartKind.ColumnClustered)]
    [InlineData(true, OfficeChartKind.ColumnClustered)]
    [InlineData(false, OfficeChartKind.Bubble)]
    [InlineData(true, OfficeChartKind.Bubble)]
    public void NativeUpdate_PreservesUnprojectablePointAppearance(bool marker, OfficeChartKind kind) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var initial = kind == OfficeChartKind.Bubble ? OfficeChartSeries.CreateBubble("Values", new[] { 1d }, new[] { 1d }, new[] { 5d }) : new OfficeChartSeries("Values", new[] { 1d });
        var chart = presentation.AddSlide().AddChart(kind, new OfficeChartData(new[] { "A" }, new[] { initial }));
        var series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants().OfType<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(item => item.LocalName == "ser");
        var point = new C.DataPoint(new C.Index { Val = 0 });
        if (marker) point.AddChild(new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Diamond }), true);
        else point.AddChild(new C.ChartShapeProperties(new A.GradientFill(new A.GradientStopList(
            new A.GradientStop(new A.RgbColorModelHex { Val = "FF0000" }) { Position = 0 },
            new A.GradientStop(new A.RgbColorModelHex { Val = "0000FF" }) { Position = 100000 }))), true);
        series.AddChild(point, true);
        string appearance = point.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        var updated = kind == OfficeChartKind.Bubble ? OfficeChartSeries.CreateBubble("Values", new[] { 1d }, new[] { 2d }, new[] { 10d }) : new OfficeChartSeries("Values", new[] { 2d });
        chart.UpdateData(new OfficeChartData(new[] { "A" }, new[] { updated }));
        series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants().OfType<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(item => item.LocalName == "ser");
        Assert.Equal(appearance, series.GetFirstChild<C.DataPoint>()!.OuterXml);
        Assert.Equal("2", ((DocumentFormat.OpenXml.OpenXmlElement?)series.GetFirstChild<C.Values>() ?? series.GetFirstChild<C.YValues>())!.Descendants<C.NumericValue>().Single().Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_IgnoresDormantMarkerOnlyLineDash(bool cap) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.Scatter, new OfficeChartData(new[] { "1" }, new[] {
            new OfficeChartSeries("Values", new[] { 2d }, new[] { 1d }, null, null, true, connectLine: false) }));
        var series = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ScatterChartSeries>().Single();
        var properties = series.GetFirstChild<C.ChartShapeProperties>()!;
        var outline = properties.GetFirstChild<A.Outline>()!;
        outline.AddChild(new A.PresetDash { Val = A.PresetLineDashValues.SystemDash }, true);
        if (cap) outline.CapType = A.LineCapValues.Round;
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
