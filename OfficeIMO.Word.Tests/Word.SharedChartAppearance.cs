using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordSharedChartAppearanceTests {
    [Theory]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Scatter)]
    public void SharedAppearance_AddsPropertiesBeforePreservedTrendlineAndErrorBars(OfficeChartType kind) {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(kind,Data(kind));
        var series = Series(chart);
        series.RemoveAllChildren<C.ChartShapeProperties>();
        series.RemoveAllChildren<C.Marker>();
        series.AddChild(new C.Trendline(new C.TrendlineType { Val = C.TrendlineValues.Linear }), true);
        series.AddChild(new C.ErrorBars(new C.ErrorBarType { Val = C.ErrorBarValues.Both },
            new C.ErrorBarValueType { Val = C.ErrorValues.FixedValue }, new C.ErrorBarValue { Val = 1 }), true);
        Assert.Empty(document.ValidateDocument());
        chart.SetData(kind, Data(kind, color: OfficeColor.Parse("#168A56")));
        Assert.Single(Series(chart).Elements<C.Trendline>());
        Assert.Single(Series(chart).Elements<C.ErrorBars>());
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void SharedAppearance_RadialOverrideHidesEveryLegendCategory(OfficeChartType renderKind) {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.Line,new OfficeChartData(new[] { "A", "B", "C" },
            new[] { new OfficeSeries("Status", new[] { 1d, 2d, 3d }, null, null, null, true,
                showInLegend: false, renderKind: renderKind) }));
        Assert.Equal(new uint[] { 0, 1, 2 }, chart.ChartPart!.ChartSpace!.Descendants<C.LegendEntry>().Select(e => e.Index!.Val!.Value));
        chart.SetData(OfficeChartType.Line, new OfficeChartData(new[] { "A", "B", "C" },
            new[] { new OfficeSeries("Updated", new[] { 3d, 2d, 1d }, null, null, null, true,
                showInLegend: false, renderKind: renderKind) }));
        Assert.Equal(new uint[] { 0, 1, 2 }, chart.ChartPart.ChartSpace.Descendants<C.LegendEntry>().Select(e => e.Index!.Val!.Value));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Radar)]
    [InlineData(OfficeChartType.Scatter)]
    public void SharedAppearance_ReconnectsLineWithoutExplicitColorOrWidth(OfficeChartType kind) {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(kind,Data(kind, connect: false));
        Assert.NotNull(Series(chart).GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>());
        chart.SetData(kind, Data(kind, connect: true));
        Assert.Null(Series(chart).GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>());
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Radar)]
    [InlineData(OfficeChartType.Scatter)]
    public void SharedAppearance_ReplacesImportedMarkerFillAndKeepsOutlineSchemaOrder(OfficeChartType kind) {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(kind,Data(kind));
        var marker = Series(chart).GetFirstChild<C.Marker>()!;
        marker.AddChild(new C.ChartShapeProperties(new A.Transform2D(), new A.NoFill(), new A.Outline(new A.NoFill(),
            new A.PresetDash { Val = A.PresetLineDashValues.Dash })), true);
        Assert.Empty(document.ValidateDocument());
        chart.SetData(kind, Data(kind, color: OfficeColor.Parse("#168A56")));
        marker = Series(chart).GetFirstChild<C.Marker>()!;
        var props = marker.ChartShapeProperties!;
        Assert.Null(props.GetFirstChild<A.NoFill>());
        Assert.Single(props.Elements<A.SolidFill>());
        Assert.Null(props.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>());
        Assert.Empty(document.ValidateDocument());
        using var bytes = new MemoryStream();
        document.Save(bytes); bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartType.Line)]
    [InlineData(OfficeChartType.Scatter)]
    public void SharedAppearance_DataUpdatePreservesExplicitMarkerFillWhenColorIsUnspecified(OfficeChartType kind) {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(kind,Data(kind));
        var series = Series(chart);
        var shape = series.GetFirstChild<C.ChartShapeProperties>()!;
        shape.GetFirstChild<A.Outline>()!.RemoveAllChildren<A.SolidFill>();
        shape.GetFirstChild<A.Outline>()!.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }), true);
        var marker = series.GetFirstChild<C.Marker>()!;
        marker.ChartShapeProperties!.RemoveAllChildren<A.SolidFill>();
        marker.ChartShapeProperties.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }), true);

        chart.SetData(kind, Data(kind));

        var updated = Series(chart).GetFirstChild<C.Marker>()!.ChartShapeProperties!;
        Assert.Equal("FFFFFF", updated.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void SharedAppearance_HiddenRadialSeriesHidesEveryCategoryAndUpdatesVisibility(OfficeChartType kind) {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(kind,Data(kind, legend: false));
        var legend = chart.ChartPart!.ChartSpace!.Descendants<C.Legend>().Single();
        Assert.Equal(new uint[] { 0, 1, 2 }, legend.Elements<C.LegendEntry>().Select(e => e.Index!.Val!.Value));
        chart.SetData(kind, Data(kind, legend: true));
        Assert.Empty(chart.ChartPart.ChartSpace.Descendants<C.LegendEntry>());
        chart.SetData(kind, Data(kind, legend: false));
        Assert.Equal(3, chart.ChartPart.ChartSpace.Descendants<C.LegendEntry>().Count());
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedAppearance_NativeMarkerSizeClampsAtBothBounds() {
        using WordDocument document = WordDocument.Create();
        foreach ((OfficeLength size, byte expected) in new[] {
            (OfficeLength.Points(1), (byte)2), (OfficeLength.Points(100), (byte)72),
            (OfficeLength.Points(8.5), (byte)9), (OfficeLength.Points(9.5), (byte)10) }) {
            var chart = document.AddChart(OfficeChartType.Line,Data(OfficeChartType.Line, markerSize: size));
            Assert.Equal(expected, Series(chart).GetFirstChild<C.Marker>()!.Size!.Val!.Value);
        }
        Assert.Empty(document.ValidateDocument());
    }

    private static DocumentFormat.OpenXml.OpenXmlCompositeElement Series(WordChart chart) =>
        chart.ChartPart!.ChartSpace!.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(e => e.LocalName == "ser");

    private static OfficeChartData Data(OfficeChartType kind, bool connect = true, bool legend = true, OfficeColor? color = null, OfficeLength? markerSize = null) =>
        new OfficeChartData(new[] { "1", "2", "3" }, new[] { new OfficeSeries("Status", new[] { 1d, 2d, 3d },
            kind == OfficeChartType.Scatter ? new[] { 1d, 2d, 3d } : null, color, null, true, legend, connect, markerSize) });
}
