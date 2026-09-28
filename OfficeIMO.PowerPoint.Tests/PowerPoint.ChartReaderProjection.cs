using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartReaderProjectionTests {
    [Fact]
    public void Snapshot_HidesLegendByPlottedOrdinalWhenNativeSeriesIndexesDiffer() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("First", new[] { 1d }), new OfficeChartSeries("Second", new[] { 2d }) }));
        C.Chart native = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!;
        native.Descendants<C.LineChartSeries>().First().Index!.Val = 5;
        C.Legend legend = native.GetFirstChild<C.Legend>()!;
        legend.AddChild(new C.LegendEntry(new C.Index { Val = 0 }, new C.Delete { Val = true }), true);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.False(snapshot.Data.Series[0].ShowInLegend);
        Assert.True(snapshot.Data.Series[1].ShowInLegend);
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Radar)]
    public void Snapshot_UsesNativeDefaultForExplicitMarkersWithoutSize(OfficeChartKind kind) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", new[] { 1d, 2d }, null, null, null, showMarkers: true) }));
        C.Marker marker = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!
            .Descendants<C.Marker>().First();
        marker.GetFirstChild<C.Size>()?.Remove();
        Assert.NotNull(marker.Symbol);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(5, snapshot.Data.Series.Single().MarkerSize);
    }

    [Fact]
    public void Snapshot_RejectsDistinctCategoryAxesInOneAxisGroup() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Columns", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered),
                new OfficeChartSeries("Line", new[] { 3d, 4d }, null, null, null, true, renderKind: OfficeChartKind.Line) }));
        C.PlotArea plot = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!
            .GetFirstChild<C.Chart>()!.PlotArea!;
        C.CategoryAxis original = plot.GetFirstChild<C.CategoryAxis>()!;
        C.CategoryAxis second = (C.CategoryAxis)original.CloneNode(true);
        second.AxisId!.Val = 700001;
        plot.GetFirstChild<C.LineChart>()!.GetFirstChild<C.AxisId>()!.Val = 700001;
        plot.Append(second);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Radar)]
    [InlineData(OfficeChartKind.Scatter)]
    public void Snapshot_RejectsPictureMarkers(OfficeChartKind kind) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(kind, new OfficeChartData(new[] { "1", "2" }, new[] {
            new OfficeChartSeries("Values", new[] { 1d, 2d }, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null) }));
        var marker = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.Marker>().Single();
        marker.Symbol!.Val = C.MarkerStyleValues.Picture;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_InheritsRadarMarkerVisibility(bool markers) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Radar, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }) }));
        var layer = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.RadarChart>().Single();
        layer.RadarStyle!.Val = markers ? C.RadarStyleValues.Marker : C.RadarStyleValues.Standard;
        layer.Elements<C.RadarChartSeries>().Single().GetFirstChild<C.Marker>()!.Remove();
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(markers, snapshot.Data.Series.Single().ShowMarkers);
    }
    [Fact]
    public void MixedSnapshot_RejectsAggregatePaddingExpansionAcrossLayers() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Long", new[] { 1d }), new OfficeChartSeries("Short", new[] { 2d }) }));
        var plot = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var original = plot.GetFirstChild<C.LineChart>()!;
        original.Elements<C.LineChartSeries>().First().GetFirstChild<C.Values>()!.Descendants<C.PointCount>().Single().Val = 25000;
        for (uint layerIndex = 1; layerIndex < 3; layerIndex++) {
            var copy = (C.LineChart)original.CloneNode(true);
            uint position = layerIndex * 2;
            foreach (var series in copy.Elements<C.LineChartSeries>()) { series.Index!.Val = position; series.Order!.Val = position++; }
            plot.InsertBefore(copy, original);
        }
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedSnapshot_RejectsIncompatibleCategoryCachesWithoutDroppingSeries(bool differentLength) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Columns", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered),
            new OfficeChartSeries("Line", new[] { 3d, 4d }, null, null, null, true, renderKind: OfficeChartKind.Line) }));
        var series = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.LineChartSeries>().Single();
        if (differentLength) {
            series.GetFirstChild<C.Values>()!.Descendants<C.PointCount>().Single().Val = 3;
            series.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.PointCount>().Single().Val = 3;
        } else series.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringPoint>().Last().NumericValue!.Text = "Different";
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MixedSnapshot_UsesTheRealCategoryCacheWhenAnotherLayerOmitsIt(bool firstLayerMissing) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "North", "South" }, new[] {
                new OfficeChartSeries("Columns", new[] { 1d, 2d }, null, null, null, true,
                    renderKind: OfficeChartKind.ColumnClustered),
                new OfficeChartSeries("Line", new[] { 3d, 4d }, null, null, null, true,
                    renderKind: OfficeChartKind.Line)
            }));
        C.ChartSpace native = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!;
        if (firstLayerMissing)
            native.Descendants<C.BarChartSeries>().Single().GetFirstChild<C.CategoryAxisData>()!.Remove();
        else
            native.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.CategoryAxisData>()!.Remove();

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { "North", "South" }, snapshot.Data.Categories);
        Assert.Equal(2, snapshot.Data.Series.Count);
    }

    [Fact]
    public void Snapshot_UsesLaterCategoryCacheWithoutTruncatingSeries() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Short", new[] { 1d, 2d, 3d }), new OfficeChartSeries("Long", new[] { 4d, 5d, 6d }) }));
        var series = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.LineChartSeries>().ToArray();
        series[0].GetFirstChild<C.CategoryAxisData>()!.Remove();
        var cache = series[0].GetFirstChild<C.Values>()!.Descendants<C.NumberingCache>().Single();
        cache.PointCount!.Val = 1;
        foreach (var point in cache.Elements<C.NumericPoint>().Skip(1).ToArray()) point.Remove();
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { "A", "B", "C" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 1d, 0d, 0d }, snapshot.Data.Series[0].Values);
        Assert.Equal(new[] { 4d, 5d, 6d }, snapshot.Data.Series[1].Values);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_RejectsExplicitMarkerFillWithUnresolvedConnectingLine(bool markersVisible) {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Line", new[] { 1d, 2d }, null, OfficeColor.Parse("#224466"), null, markersVisible) }));
        var series = document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.LineChartSeries>().Single();
        series.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!.GetFirstChild<A.SolidFill>()!.Remove();
        series.GetFirstChild<C.Marker>()!.ChartShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val = "00FF00";
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void MixedSnapshot_ValidatesWholePlotBudgetBeforeReadingIndividualLayers() {
        using var document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Columns", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered),
            new OfficeChartSeries("Line", new[] { 3d, 4d }, null, null, null, true, renderKind: OfficeChartKind.Line) }));
        foreach (var cache in document.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.NumberingCache>()) cache.PointCount!.Val = 60000;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }
}
