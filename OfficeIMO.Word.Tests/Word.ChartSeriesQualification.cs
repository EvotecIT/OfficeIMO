using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public class WordChartSeriesQualificationTests {
    [Theory]
    [InlineData(WordChartBarGrouping.Stacked)]
    [InlineData(WordChartBarGrouping.PercentStacked)]
    public void AuthoredStackedBarsKeepOverlapAndRemainProjectable(WordChartBarGrouping grouping) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart();
        chart.AddCategories(new System.Collections.Generic.List<string> { "A", "B" });
        chart.AddBar("Results", new[] { 1, 2 }, OfficeColor.Blue);
        chart.BarGrouping = grouping;
        C.BarChart native = chart.ChartPart!.ChartSpace!.Descendants<C.BarChart>().Single();
        Assert.Equal(100, native.GetFirstChild<C.Overlap>()!.Val!.Value);
        Assert.True(chart.TryGetSnapshot(out _));
        chart.BarGrouping = WordChartBarGrouping.Clustered;
        Assert.Equal(0, native.GetFirstChild<C.Overlap>()!.Val!.Value);
    }

    [Fact]
    public void OfficeSnapshot_HidesLegendByPlottedOrdinal() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("First", new[] { 1d }), new OfficeChartSeries("Second", new[] { 2d }) }));
        C.Chart native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        native.Descendants<C.LineChartSeries>().First().Index!.Val = 5;
        native.GetFirstChild<C.Legend>()!.AddChild(
            new C.LegendEntry(new C.Index { Val = 0 }, new C.Delete { Val = true }), true);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.False(snapshot.Data.Series[0].ShowInLegend);
        Assert.True(snapshot.Data.Series[1].ShowInLegend);
    }

    [Fact]
    public void OfficeSnapshot_RejectsDistinctCategoryAxesInOneGroup() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Columns", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.ColumnClustered),
                new OfficeChartSeries("Line", new[] { 3d, 4d }, null, null, null, true, renderKind: OfficeChartKind.Line) }));
        C.PlotArea plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        C.CategoryAxis second = (C.CategoryAxis)plot.GetFirstChild<C.CategoryAxis>()!.CloneNode(true);
        second.AxisId!.Val = 700001;
        plot.GetFirstChild<C.LineChart>()!.GetFirstChild<C.AxisId>()!.Val = 700001;
        plot.Append(second);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UpdatingAreaSeriesCanRestoreAnInheritedOutline(bool zeroWidth) {
        OfficeChartData Data(bool connectLine) => new(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Area", new[] { 1d, 2d }, null, null, null,
                showMarkers: false, connectLine: connectLine)
        });
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Area, Data(false));
        A.Outline Outline() => chart.ChartPart!.ChartSpace!.Descendants<C.AreaChartSeries>().Single()
            .GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!;
        Assert.NotNull(Outline().GetFirstChild<A.NoFill>());
        if (zeroWidth) Outline().Width = 0;
        chart.SetData(OfficeChartKind.Area, Data(true));
        Assert.Null(Outline().GetFirstChild<A.NoFill>());
        Assert.NotEqual(0, Outline().Width?.Value);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.True(snapshot.Data.Series[0].ConnectLine);
    }

    [Theory]
    [InlineData(OfficeChartKind.Area)]
    [InlineData(OfficeChartKind.AreaStacked)]
    [InlineData(OfficeChartKind.AreaStacked100)]
    public void Snapshot_MapsHiddenAreaOutlinesToFillOnlyGeometry(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", new[] { 1d, 2d }, null, OfficeColor.Parse("#123456")) }));
        var outline = chart.ChartPart!.ChartSpace!.Descendants<C.AreaChartSeries>().Single()
            .GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!;
        outline.RemoveAllChildren(); outline.Append(new A.NoFill());
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.False(Assert.Single(snapshot.Data.Series).ConnectLine);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
        var area = Assert.Single(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Polygon &&
            shape.Shape.FillColor == OfficeColor.Parse("#123456"));
        Assert.Null(area.Shape.StrokeColor);
    }
    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Scatter)]
    public void Snapshot_RejectsPerPointPictureMarkers(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(kind, new OfficeChartData(new[] { "1", "2" }, new[] {
            new OfficeChartSeries("Values", new[] { 1d, 2d }, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null) }));
        var series = chart.ChartPart!.ChartSpace!.Descendants<OpenXmlCompositeElement>().Single(item => item.LocalName == "ser");
        series.InsertBefore(new C.DataPoint(new C.Index { Val = 0 },
            new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Picture })), (OpenXmlElement?)series.GetFirstChild<C.CategoryAxisData>() ?? series.GetFirstChild<C.XValues>());
        Assert.False(chart.TryGetSnapshot(out _));
    }
    [Fact]
    public void Snapshot_PreservesFilledSeriesOutlineAndPointOverrides() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }, null, OfficeColor.Parse("#123456")) }));
        var series = chart.ChartPart!.ChartSpace!.Descendants<C.BarChartSeries>().Single();
        var outline = series.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!;
        outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val = "ABCDEF";
        outline.Width = 25400;
        series.InsertBefore(new C.DataPoint(new C.Index { Val = 1 },
            new C.ChartShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()))), series.GetFirstChild<C.Values>());
        Assert.True(chart.TryGetSnapshot(out var snapshot));
        var styles = snapshot.Data.Series.Single().ToOfficeSeries().PointStyles!;
        Assert.Equal(OfficeColor.Parse("#ABCDEF"), styles[0]!.OutlineColor);
        Assert.Equal(2d, styles[0]!.OutlineWidth);
        Assert.True(styles[1]!.NoFill);
        Assert.False(styles[1]!.ShowOutline);
    }
    [Theory]
    [InlineData(OfficeChartKind.ColumnStacked)]
    [InlineData(OfficeChartKind.Scatter)]
    public void Snapshot_PreservesNativePlottingOrder(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(kind, new OfficeChartData(new[] { "1" }, new[] {
            new OfficeChartSeries("First", new[] { 10d }, kind == OfficeChartKind.Scatter ? new[] { 1d } : null),
            new OfficeChartSeries("Second", new[] { 20d }, kind == OfficeChartKind.Scatter ? new[] { 2d } : null) }));
        var orders = chart.ChartPart!.ChartSpace!.Descendants<C.Order>().ToArray();
        orders[0].Val = 1; orders[1].Val = 0;
        Assert.True(chart.TryGetSnapshot(out var snapshot));
        Assert.Equal(new[] { "Second", "First" }, snapshot.Data.Series.Select(s => s.Name));
    }

    [Theory]
    [InlineData(false, "gradient")]
    [InlineData(false, "noFill")]
    [InlineData(false, "compound")]
    [InlineData(false, "customDash")]
    [InlineData(true, "gradient")]
    [InlineData(true, "compound")]
    [InlineData(true, "customDash")]
    [InlineData(true, "marker")]
    [InlineData(true, "miterLimit")]
    public void Snapshot_RejectsUnrepresentedNativeAppearance(bool point, string appearance) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(appearance == "marker" ? OfficeChartKind.Line : OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 1d }) }));
        var series = chart.ChartPart!.ChartSpace!.Descendants<OpenXmlCompositeElement>().Single(item => item.LocalName == "ser");
        var properties = new C.ChartShapeProperties();
        if (appearance == "gradient") properties.Append(new A.GradientFill());
        if (appearance == "noFill") properties.Append(new A.NoFill());
        if (appearance == "compound") properties.Append(new A.Outline { CompoundLineType = A.CompoundLineValues.Double });
        if (appearance == "customDash") properties.Append(new A.Outline(new A.CustomDash()));
        if (appearance == "miterLimit") properties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "445566" }), new A.Miter { Limit = 500000 }));
        if (point) {
            var item = new C.DataPoint(new C.Index { Val = 0 });
            if (appearance == "marker") item.Append(new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Picture }));
            else item.Append(properties);
            series.InsertBefore(item, series.GetFirstChild<C.Values>());
        } else {
            series.GetFirstChild<C.ChartShapeProperties>()?.Remove();
            series.AddChild(properties, true);
        }
        string before = chart.ChartPart.ChartSpace.OuterXml;
        Assert.False(chart.TryGetSnapshot(out _));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
    }
}
