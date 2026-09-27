using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartAppearanceIntegrityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedReader_QualifiesEffectiveGroupSmoothingBeforeProjection(bool overrideWithStraightLine) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Line", new[] { 1d, 2d }) }));
        var layer = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.LineChart>().Single();
        layer.AddChild(new C.Smooth { Val = true }, true);
        var series = layer.Elements<C.LineChartSeries>().Single();
        series.GetFirstChild<C.Smooth>()?.Remove();
        if (overrideWithStraightLine) series.AddChild(new C.Smooth { Val = false }, true);
        using var bytes = new MemoryStream(document.ToBytes());
        using var reopened = PowerPointPresentation.Load(bytes);
        Assert.Equal(overrideWithStraightLine, reopened.Slides.Single().Charts.Single().TryGetOfficeSnapshot(out _));
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SharedReader_InheritsGroupMarkerAndConnectionVisibilityAcrossSaveReopen(int mode) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(mode == 0 ? OfficeChartKind.Line : OfficeChartKind.Scatter,
            new OfficeChartData(new[] { "1", "2" }, new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }, mode == 0 ? null : new[] { 1d, 2d }) }));
        var plot = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        if (mode == 0) plot.GetFirstChild<C.LineChart>()!.AddChild(new C.ShowMarker { Val = false }, true);
        else plot.GetFirstChild<C.ScatterChart>()!.ScatterStyle!.Val = mode == 1 ? C.ScatterStyleValues.Line : C.ScatterStyleValues.Marker;
        plot.Descendants<C.Marker>().Single().Remove();
        using var bytes = new MemoryStream(document.ToBytes());
        using var reopened = PowerPointPresentation.Load(bytes);
        Assert.True(reopened.Slides.Single().Charts.Single().TryGetOfficeSnapshot(out var snapshot));
        var series = snapshot.Data.Series.Single();
        Assert.Equal(mode == 2, series.ShowMarkers);
        Assert.Equal(mode != 2, series.ConnectLine);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void SharedReader_RejectsTheAggregateCacheBudgetAcrossMixedLayers() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Line", new[] { 1d, 2d }),
            new OfficeChartSeries("Columns", new[] { 3d, 4d }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered) }));
        var plot = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        foreach (var count in plot.Descendants<C.PointCount>()) count.Val = 50000;
        Assert.True(chart.TryGetOfficeSnapshot(out var bounded));
        Assert.Equal(2, bounded.Data.Series.Count);
        plot.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.Values>()!.Descendants<C.PointCount>().Single().Val = 50001;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedReader_ReportsUnmappedDashWhileKeepingNativeDataUpdatesAvailable() {
        using PowerPointPresentation document = PowerPointPresentation.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 1d, 2d },
            null, OfficeColor.Parse("#168A56")) });
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, data);
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        var outline = part.ChartSpace!.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!;
        outline.AddChild(new A.PresetDash { Val = A.PresetLineDashValues.LargeDash }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        chart.UpdateData(data);
        Assert.Equal(A.PresetLineDashValues.LargeDash,
            part.ChartSpace.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!.GetFirstChild<A.PresetDash>()!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedReader_PreservesMarkerAppearanceAndSeriesDash() {
        using PowerPointPresentation document = PowerPointPresentation.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 1d, 2d },
            null, OfficeColor.Parse("#168A56"), null, true, markerSize: 9, markerShape: OfficeChartMarkerShape.Diamond,
            markerOutlineColor: OfficeColor.Parse("#333333"), markerOutlineWidth: 2, strokeWidth: 3,
            strokeDashStyle: OfficeStrokeDashStyle.Dash) });
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, data);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        var series = snapshot.Data.Series.Single();
        Assert.Equal(9, series.MarkerSize);
        Assert.Equal(OfficeChartMarkerShape.Diamond, series.MarkerShape);
        Assert.Equal(OfficeColor.Parse("#333333"), series.MarkerOutlineColor);
        Assert.Equal(2d, series.MarkerOutlineWidth);
        Assert.Equal(3d, series.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dash, series.StrokeDashStyle);
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Radar)]
    [InlineData(OfficeChartKind.Scatter)]
    public void SharedAppearance_PowerPointRestoresLineAndReplacesImportedMarkerFill(OfficeChartKind kind) {
        OfficeChartData Data(bool connect, OfficeColor? color = null) => new OfficeChartData(new[] { "1", "2" },
            new[] { new OfficeChartSeries("Status", new[] { 1d, 2d }, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null,
                color, null, true, connectLine: connect, markerSize: 1) });
        using PowerPointPresentation document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(kind, Data(false));
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot disabled));
        Assert.False(disabled.Data.Series.Single().ConnectLine);
        chart.UpdateData(Data(true));
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot enabled));
        Assert.True(enabled.Data.Series.Single().ConnectLine);
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        var series = part.ChartSpace!.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(e => e.LocalName == "ser");
        Assert.Null(series.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>());
        series.GetFirstChild<C.Marker>()!.AddChild(new C.ChartShapeProperties(new A.NoFill()), true);
        chart.UpdateData(Data(true, OfficeColor.Parse("#168A56")));
        series = part.ChartSpace.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(e => e.LocalName == "ser");
        Assert.Null(series.GetFirstChild<C.Marker>()!.ChartShapeProperties!.GetFirstChild<A.NoFill>());
        Assert.Equal((byte)2, series.GetFirstChild<C.Marker>()!.Size!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
        using var bytes = new MemoryStream(document.ToBytes());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        Assert.Empty(reopened.ValidateDocument());
    }
}
