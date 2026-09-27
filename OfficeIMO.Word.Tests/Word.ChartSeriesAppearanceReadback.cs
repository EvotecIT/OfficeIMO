using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartSeriesAppearanceReadbackTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Snapshot_QualifiesEffectiveGroupSmoothingWithSeriesOverride(bool overrideWithStraightLine, bool explicitGroupValue) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Line", new[] { 1d, 2d }) }));
        var layer = chart.ChartPart!.ChartSpace!.Descendants<C.LineChart>().Single();
        layer.AddChild(explicitGroupValue ? new C.Smooth { Val = true } : new C.Smooth(), true);
        var series = layer.Elements<C.LineChartSeries>().Single();
        series.GetFirstChild<C.Smooth>()?.Remove();
        if (overrideWithStraightLine) series.AddChild(new C.Smooth { Val = false }, true);
        Assert.Equal(overrideWithStraightLine, chart.TryGetSnapshot(out _));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_InheritsNativeGroupMarkerAndConnectionVisibility(bool scatter) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(scatter ? OfficeChartKind.Scatter : OfficeChartKind.Line,
            new OfficeChartData(new[] { "1", "2" }, new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }, scatter ? new[] { 1d, 2d } : null) }));
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        if (scatter) {
            var layer = plot.GetFirstChild<C.ScatterChart>()!;
            layer.ScatterStyle!.Val = C.ScatterStyleValues.Marker;
            layer.Elements<C.ScatterChartSeries>().Single().GetFirstChild<C.Marker>()?.Remove();
        } else {
            var layer = plot.GetFirstChild<C.LineChart>()!;
            layer.AddChild(new C.ShowMarker { Val = false }, true);
            layer.Elements<C.LineChartSeries>().Single().GetFirstChild<C.Marker>()?.Remove();
        }
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = WordDocument.Load(bytes);
        Assert.True(reopened.Charts.Single().TryGetSnapshot(out var snapshot));
        var series = snapshot.Data.Series.Single().ToOfficeSeries();
        Assert.Equal(scatter, series.ShowMarkers);
        Assert.Equal(!scatter, series.ConnectLine);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Snapshot_RejectsUnqualifiedAppearanceAndAllowsAnExplicitStyleReplacement(int appearance) {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "1", "2" }, new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }, null,
            OfficeColor.Parse("#224466"), null, true, markerSize: 9, strokeWidth: 2) });
        var chart = document.AddChart(OfficeChartKind.Line, data);
        var native = chart.ChartPart!.ChartSpace!.Descendants<C.LineChartSeries>().Single();
        if (appearance == 0) native.AddChild(new C.Smooth { Val = true }, true);
        else {
            var markerProperties = native.GetFirstChild<C.Marker>()!.ChartShapeProperties!;
            if (appearance == 1) markerProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val = "FF0000";
            else markerProperties.AddChild(new A.Outline(new A.NoFill()), true);
        }
        Assert.False(chart.TryGetSnapshot(out _));
        chart.SetData(OfficeChartKind.Line, data);
        Assert.Equal(appearance != 0, chart.TryGetSnapshot(out _));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_RejectsAggregateSparseCacheBudgetBeforeProjection(bool scatter) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(scatter ? OfficeChartKind.Scatter : OfficeChartKind.Line,
            new OfficeChartData(new[] { "1", "2" }, new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }, scatter ? new[] { 1d, 2d } : null) }));
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var layer = (DocumentFormat.OpenXml.OpenXmlCompositeElement)plot.ChildElements.First(element => element.LocalName.EndsWith("Chart"));
        var original = (DocumentFormat.OpenXml.OpenXmlCompositeElement)layer.ChildElements.First(element => element.LocalName == "ser");
        foreach (var count in original.Descendants<C.PointCount>()) count.Val = 10000;
        for (uint index = 1; index < 10; index++) {
            var copy = (DocumentFormat.OpenXml.OpenXmlCompositeElement)original.CloneNode(true);
            copy.GetFirstChild<C.Index>()!.Val = index; copy.GetFirstChild<C.Order>()!.Val = index;
            layer.InsertAfter(copy, original);
        }
        Assert.True(chart.TryGetSnapshot(out var atLimit));
        Assert.Equal(10, atLimit.Data.Series.Count);
        var excess = (DocumentFormat.OpenXml.OpenXmlCompositeElement)original.CloneNode(true);
        excess.GetFirstChild<C.Index>()!.Val = 10; excess.GetFirstChild<C.Order>()!.Val = 10;
        layer.InsertAfter(excess, original);
        Assert.False(chart.TryGetSnapshot(out _));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Scatter)]
    public void Snapshot_RetainsNativeLineAndMarkerAppearanceAcrossSaveReopen(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var expected = new OfficeChartSeries("Styled", new[] { 1d, 2d }, new[] { 1d, 2d }, OfficeColor.Parse("#224466"), null, true,
            connectLine: false, markerSize: 9, markerShape: OfficeChartMarkerShape.Diamond,
            markerOutlineColor: OfficeColor.Parse("#FF0000"), markerOutlineWidth: 2, strokeWidth: 3, strokeDashStyle: OfficeStrokeDashStyle.Dash);
        var chart = document.AddChart(kind, new OfficeChartData(new[] { "1", "2" }, new[] { expected }));
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = WordDocument.Load(bytes);
        Assert.True(reopened.Charts.Single().TryGetSnapshot(out var snapshot));
        var series = snapshot.Data.Series.Single().ToOfficeSeries();
        Assert.False(series.ConnectLine);
        Assert.Equal(9, series.MarkerSize);
        Assert.Equal(OfficeChartMarkerShape.Diamond, series.MarkerShape);
        Assert.Equal(expected.MarkerOutlineColor, series.MarkerOutlineColor);
        Assert.Equal(2d, series.MarkerOutlineWidth);
        Assert.Equal(3d, series.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dash, series.StrokeDashStyle);
        Assert.Equal(expected.Color, series.Color);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void Snapshot_RejectsAnUnqualifiedNativeDashInsteadOfFlatteningItsAppearance() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Status", new[] { 1d, 2d }, null, null, null, true, strokeWidth: 2) }));
        chart.ChartPart!.ChartSpace!.Descendants<C.LineChartSeries>().Single().ChartShapeProperties!
            .GetFirstChild<A.Outline>()!.AddChild(new A.PresetDash { Val = A.PresetLineDashValues.LargeDash }, true);
        Assert.False(chart.TryGetSnapshot(out _));
        Assert.Empty(document.ValidateDocument());
    }
}
