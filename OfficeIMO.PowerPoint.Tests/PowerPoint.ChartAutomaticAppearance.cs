using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

/// <summary>Pins automatic chart appearance to Microsoft Office exports.</summary>
public sealed class PowerPointChartAutomaticAppearanceTests {
    private static string FixtureFolder => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets",
        "Documents", "Charts", "AutomaticColors");

    private static string RoadmapPath => Path.Combine(RepositoryTestPaths.Find(), "Assets", "PowerPointTemplates",
        "PowerPointAdvancedRoadmap.pptx");

    public static IEnumerable<object[]> ReferenceCharts() => ReadReference().Select(chart => new object[] { chart.Number });

    [Theory]
    [MemberData(nameof(ReferenceCharts))]
    public void AutomaticColorsMatchThePowerPointExport(int number) {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(Path.Combine(FixtureFolder, "automatic-colors.pptx"));
        PowerPointChart chart = presentation.Slides[number - 1].Charts.Single();

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot, out string? reason), reason);
        OfficeChartSnapshot office = PowerPointChartSnapshotMapper.ToOfficeSnapshot(snapshot, snapshot.WidthPoints, snapshot.HeightPoints);

        Assert.Equal(ReadReference().Single(chart => chart.Number == number).Fills, DrawnFills(office));
    }

    [Fact]
    public void AutomaticLineSeriesUseOfficeMarkerSequence() {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(Path.Combine(FixtureFolder, "automatic-colors.pptx"));
        PowerPointChart chart = presentation.Slides[5].Charts.Single();

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot, out string? reason), reason);
        Assert.Equal(new OfficeChartMarkerShape?[] { OfficeChartMarkerShape.Diamond, OfficeChartMarkerShape.Square, OfficeChartMarkerShape.Triangle },
            snapshot.Data.Series.Select(series => series.SharedAppearance!.MarkerShape).ToArray());
        Assert.All(snapshot.Data.Series, series => Assert.Equal(7, series.SharedAppearance!.MarkerSize));
    }

    [Fact]
    [Trait("Category", "Performance")]
    public void AutomaticColorsProjectDensePlotSeriesWithoutRepeatedPlotScans() {
        const int count = 1500;
        var plot = new C.PlotArea();
        var bars = new C.BarChart();
        plot.Append(bars);
        for (int index = 0; index < count; index++) {
            bars.Append(new C.BarChartSeries(
                new C.Index { Val = (uint)index },
                new C.Order { Val = (uint)index },
                new C.Values(new C.NumberLiteral(
                    new C.PointCount { Val = 1U },
                    new C.NumericPoint(new C.NumericValue("1")) { Index = 0U }))));
        }

        var timer = System.Diagnostics.Stopwatch.StartNew();
        OfficeOpenXmlChartSeriesReader.Result? result = OfficeOpenXmlChartSeriesReader.ReadCategories(
            bars.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(),
            OfficeChartKind.ColumnClustered, scheme: null, OfficeChartAxisGroup.Primary,
            maximumPoints: count);
        timer.Stop();

        Assert.NotNull(result);
        Assert.Equal(count, result!.Series.Count);
        Assert.Equal((uint)(count - 1), result.Series[result.Series.Count - 1].SourceIndex);
        Console.WriteLine($"Automatic chart colours: {count} plot series, {timer.Elapsed.TotalMilliseconds:F2} ms.");
    }

    [Fact]
    public void DenseLayoutReferenceDeckDrawsItsAutomaticColumnChart() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Pdf.Tests", "Pdf", "ReferenceBaselines",
            "native-powerpoint-dense-layout.pptx");
        using PowerPointPresentation presentation = PowerPointPresentation.Load(path);
        PdfDocumentConversionResult result = presentation.ToPdfDocumentResult();

        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-chart");
        PowerPointChart chart = presentation.Slides.SelectMany(slide => slide.Charts).Single();
        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot, out _));
        Assert.Equal(OfficeColor.Parse("#156082"), snapshot.Data.Series.Single().Color);
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void FlatRibbonUsesItsBodyPaintAndRetainsAnInvisibleNativeSeries(bool fill, bool outline) {
        using var presentation = PowerPointPresentation.Load(RoadmapPath);
        var slide = presentation.Slides.First(item => item.SlidePart.ChartParts.Any(part => part.ChartSpace!.Descendants<C.Line3DChart>().Any()));
        var part = slide.SlidePart.ChartParts.First(item => item.ChartSpace!.Descendants<C.Line3DChart>().Any());
        foreach (var native in part.ChartSpace!.Descendants<C.Line3DChart>().SelectMany(layer => layer.Elements<C.LineChartSeries>())) {
            native.RemoveAllChildren<C.Marker>();
            native.AddChild(new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Circle }), true);
            var properties = native.GetFirstChild<C.ChartShapeProperties>()!;
            properties.RemoveAllChildren<A.SolidFill>(); properties.RemoveAllChildren<A.NoFill>();
            properties.AddChild(fill ? (DocumentFormat.OpenXml.OpenXmlElement)new A.SolidFill(new A.RgbColorModelHex { Val = "0000FF" }) : new A.NoFill(), true);
            var line = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            line.RemoveAllChildren<A.SolidFill>(); line.RemoveAllChildren<A.NoFill>(); line.Width = 38100;
            line.AddChild(outline ? (DocumentFormat.OpenXml.OpenXmlElement)new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }) : new A.NoFill(), true);
            if (line.Parent == null) properties.AddChild(line, true);
        }
        var chart = slide.Charts.Single(item => item.InspectImportedContent().Family == "line3DChart");
        Assert.True(chart.TryGetSnapshot(out var nativeSnapshot, out var reason), reason);
        var snapshot = PowerPointChartSnapshotMapper.ToOfficeSnapshot(nativeSnapshot, 640, 360);
        Assert.All(snapshot.Data.Series, series => {
            Assert.False(series.ShowMarkers);
            Assert.Equal(!fill, series.SeriesStyle!.NoFill);
            Assert.Equal(fill || outline, series.SeriesStyle.ShowOutline);
            if (fill || outline) Assert.Equal(fill ? OfficeColor.Blue : OfficeColor.Red, series.Color);
        });
        var drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Equal(fill || outline, drawing.Shapes.Any(shape => shape.Shape.StrokeWidth == 3 &&
            shape.Shape.StrokeColor == (fill ? OfficeColor.Blue : OfficeColor.Red) &&
            shape.Shape.Kind is OfficeShapeKind.Line or OfficeShapeKind.Path));
    }

    [Fact]
    public void FlatRibbonAutomaticBodyAcceptsCallerPaletteIndependentlyOfItsNativeBorder() {
        using var presentation = PowerPointPresentation.Load(RoadmapPath);
        var slide = presentation.Slides.First(item => item.SlidePart.ChartParts.Any(part => part.ChartSpace!.Descendants<C.Line3DChart>().Any()));
        var part = slide.SlidePart.ChartParts.First(item => item.ChartSpace!.Descendants<C.Line3DChart>().Any());
        foreach (var native in part.ChartSpace!.Descendants<C.Line3DChart>().SelectMany(layer => layer.Elements<C.LineChartSeries>())) {
            var properties = native.GetFirstChild<C.ChartShapeProperties>()!;
            properties.RemoveAllChildren<A.SolidFill>(); properties.RemoveAllChildren<A.NoFill>();
            var line = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            line.RemoveAllChildren<A.SolidFill>(); line.RemoveAllChildren<A.NoFill>(); line.Width = 38100;
            line.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }), true);
            if (line.Parent == null) properties.AddChild(line, true);
        }
        var chart = slide.Charts.Single(item => item.InspectImportedContent().Family == "line3DChart");
        Assert.True(chart.TryGetSnapshot(out var nativeSnapshot, out var reason), reason);
        Assert.All(nativeSnapshot.Data.Series, series => Assert.True(series.HasAutomaticSeriesColor));
        var snapshot = PowerPointChartSnapshotMapper.ToOfficeSnapshot(nativeSnapshot, 640, 360,
            new OfficeChartStyle(palette: new[] { OfficeColor.Green }));
        Assert.Contains(OfficeChartDrawingRenderer.Render(snapshot).Shapes,
            shape => shape.Shape.StrokeWidth == 3 && shape.Shape.StrokeColor == OfficeColor.Green &&
                shape.Shape.Kind is OfficeShapeKind.Line or OfficeShapeKind.Path);
    }

    [Fact]
    public void SkippedChartWarningNamesTheUnsupportedAppearance() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Values", new[] { 2d, 3d }) }));
        ChartPart part = presentation.Slides.Single().SlidePart.ChartParts.Single();
        C.ChartShapeProperties properties = part.ChartSpace!.Descendants<C.BarChartSeries>().Single()
            .GetFirstChild<C.ChartShapeProperties>()!;
        properties.RemoveAllChildren<A.SolidFill>();
        properties.InsertAt(new A.GradientFill(new A.GradientStopList(
            new A.GradientStop(new A.RgbColorModelHex { Val = "112233" }) { Position = 0 },
            new A.GradientStop(new A.RgbColorModelHex { Val = "445566" }) { Position = 100000 })), 0);

        PdfDocumentConversionResult result = presentation.ToPdfDocumentResult();

        PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "unsupported-chart");
        Assert.Contains("'Values' uses a series fill, line or effect", warning.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MicrosoftAdvancedChartFamiliesRenderAndReportFlatProjections() {
        using PowerPointPresentation presentation = PowerPointPresentation.Load(RoadmapPath);
        PdfDocumentConversionResult result = presentation.ToPdfDocumentResult();

        Assert.DoesNotContain(result.Warnings, warning => warning.Code is "unsupported-chart" or "unsupported-smartart");
        string[] projected = result.Warnings.Where(warning => warning.Code == "chart-projected")
            .Select(warning => warning.Message.Split(' ')[2]).OrderBy(name => name, StringComparer.Ordinal).ToArray();
        Assert.Equal(new[] { "area3DChart", "bar3DChart", "line3DChart", "ofPieChart", "pie3DChart", "surface3DChart", "surfaceChart" },
            projected);

        PowerPointChart[] charts = presentation.Slides.SelectMany(slide => slide.Charts).ToArray();
        PowerPointChartSnapshot line = Snapshot(charts, "line3DChart");
        Assert.All(line.Data.Series, series => {
            Assert.True(series.SharedAppearance!.ConnectLine);
            Assert.False(series.SharedAppearance.ShowMarkers);
        });
        OfficeChartSnapshot flatRibbon = PowerPointChartSnapshotMapper.ToOfficeSnapshot(line, 640, 360);
        OfficeDrawing ribbonDrawing = OfficeChartDrawingRenderer.Render(flatRibbon);
        foreach (OfficeChartSeries projectedSeries in flatRibbon.Data.Series)
            Assert.Contains(ribbonDrawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                shape.Shape.StrokeColor == projectedSeries.Color);
        PowerPointChartSnapshot stock = Snapshot(charts, "stockChart");
        Assert.True(stock.Layout.ShowHighLowLines);
        Assert.All(Snapshot(charts, "surfaceChart").Data.Series, series => Assert.False(series.SharedAppearance!.ShowMarkers));
        PowerPointChartSnapshot area = Snapshot(charts, "area3DChart");
        Assert.Equal("m/d/yyyy", area.Layout.CategoryAxisNumberFormat);
    }

    private static PowerPointChartSnapshot Snapshot(IEnumerable<PowerPointChart> charts, string family) {
        PowerPointChart chart = charts.Single(item => item.InspectImportedContent().Family == family);
        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot, out string? reason), reason);
        return snapshot;
    }

    // The fills a chart draws in series order: filled series and visible markers per series,
    // and slices per point for radial charts.
    internal static IReadOnlyList<string> DrawnFills(OfficeChartSnapshot snapshot) {
        var fills = new List<string>();
        void Add(OfficeColor color) {
            string hex = color.ToRgbHex();
            if (!fills.Contains(hex)) fills.Add(hex);
        }
        if (snapshot.ChartKind is OfficeChartKind.Pie or OfficeChartKind.Doughnut) {
            for (int index = 0; index < snapshot.Data.Categories.Count; index++) Add(snapshot.Style.GetSeriesColor(index));
            return fills;
        }
        bool lines = snapshot.ChartKind is OfficeChartKind.Line or OfficeChartKind.Scatter;
        foreach (OfficeChartSeries series in snapshot.Data.Series) {
            if (lines && (!series.ShowMarkers || series.MarkerShape == OfficeChartMarkerShape.None)) continue;
            Add(series.Color!.Value);
        }
        return fills;
    }

    internal static IReadOnlyList<(int Number, string Title, string[] Fills)> ReadReference() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureFolder, "office-reference.json")));
        return document.RootElement.GetProperty("charts").EnumerateArray().Select(chart => (
            chart.GetProperty("number").GetInt32(),
            chart.GetProperty("title").GetString()!,
            chart.GetProperty("officeFills").EnumerateArray().Select(fill => fill.GetString()!).ToArray())).ToArray();
    }
}
