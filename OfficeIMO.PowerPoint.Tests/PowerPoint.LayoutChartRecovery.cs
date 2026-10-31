using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointLayoutChartRecoveryTests {
    private static OfficeChartData Combo() => new(new[] { "Q1", "Q2" }, new[] {
        new OfficeSeries("Adoption", new[] { 28d, 43d }, null, null, null,
            showMarkers: false, renderKind: OfficeChartType.ColumnClustered),
        new OfficeSeries("Conversion", new[] { 6.2d, 7.8d }, null, null, null,
            showMarkers: true, renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
    });

    [Fact]
    public void ComboSeriesStylesUseAnIndexAcrossChartGroups() {
        using var document = PowerPointDocument.Create();
        var chart = document.AddSlide().AddChart(OfficeChartType.ColumnClustered, Combo());
        chart.SetSeriesFillColor(0, "123456");
        chart.SetSeriesLineColor(1, "654321", OfficeLength.Points(2));
        chart.SetSeriesMarker(1, OfficeChartMarkerShape.Diamond, OfficeLength.Points(9));
        Assert.Throws<InvalidOperationException>(() => chart.SetSeriesMarker(0, OfficeChartMarkerShape.Circle));
        C.BarChartSeries bar = chart.GetChart().Descendants<C.BarChartSeries>().Single();
        C.LineChartSeries line = chart.GetChart().Descendants<C.LineChartSeries>().Single();
        Assert.Equal("123456", bar.ChartShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Equal("654321", line.ChartShapeProperties!.GetFirstChild<A.Outline>()!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Equal(C.MarkerStyleValues.Diamond, line.Marker!.Symbol!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void LayoutComboChartUsesEachSeriesPaintKindAndKeepsTheSavedDeckValid() {
        using var document = PowerPointDocument.Create();
        var content = new PowerPointLayoutSlide { Title = "Quarterly adoption" };
        content.Blocks.Add(new PowerPointLayoutChart(OfficeChartType.ColumnClustered, Combo()));
        PowerPointSlide slide = Assert.Single(PowerPointLayoutEngine.AddSlides(document, content));
        var chart = Assert.Single(slide.Charts);
        var bar = chart.GetChart().Descendants<C.BarChartSeries>().Single();
        var line = chart.GetChart().Descendants<C.LineChartSeries>().Single();
        Assert.Equal(OfficeTheme.Default.Palette.Accent1.ToRgbHex(), bar.ChartShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Equal(OfficeTheme.Default.Palette.Accent2.ToRgbHex(), line.ChartShapeProperties!.GetFirstChild<A.Outline>()!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        using var saved = new MemoryStream();
        document.Save(saved);
        saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        Assert.Single(reopened.Slides);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LayoutLineSeriesKeepsAnExplicitlyDisabledConnectingLine(bool combo) {
        using var document = PowerPointDocument.Create();
        var content = new PowerPointLayoutSlide { Title = "Independent observations" };
        var points = new OfficeSeries("Observations", new[] { 2d, 5d }, null, null, null,
            showMarkers: true, connectLine: false, renderKind: OfficeChartType.Line);
        var data = new OfficeChartData(new[] { "A", "B" }, combo
            ? new[] { new OfficeSeries("Columns", new[] { 3d, 6d }), points } : new[] { points });
        content.Blocks.Add(new PowerPointLayoutChart(combo ? OfficeChartType.ColumnClustered : OfficeChartType.Line, data));
        var slide = Assert.Single(PowerPointLayoutEngine.AddSlides(document, content));
        var line = Assert.Single(slide.Charts).GetChart().Descendants<C.LineChartSeries>().Single();
        Assert.NotNull(line.ChartShapeProperties!.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>());
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void LayoutPercentStackedLineKeepsLinePaintWhenItOverridesTheDefaultKind() {
        using var document = PowerPointDocument.Create();
        var theme = PowerPointLayoutTheme.From(OfficeTheme.Classic);
        var content = new PowerPointLayoutSlide { Title = "Percent-stacked observations" };
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Observations", new[] { 2d, 5d }, null, null, null,
                showMarkers: true, renderKind: OfficeChartType.LineStacked100)
        });
        content.Blocks.Add(new PowerPointLayoutChart(OfficeChartType.Line, data));
        var slide = Assert.Single(PowerPointLayoutEngine.AddSlides(document, content, theme));
        var line = Assert.Single(slide.Charts).GetChart().Descendants<C.LineChartSeries>().Single();
        var outline = line.ChartShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal(25400, outline.Width!.Value);
        Assert.Equal(theme.Accent1, outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LayoutAutomaticPalettePreservesAuthoredNoFillAndHiddenOutlines(bool line) {
        using var document = PowerPointDocument.Create();
        var content = new PowerPointLayoutSlide { Title = "Authored chart appearance" };
        var series = new OfficeSeries("Observations", new[] { 2d, 5d }).WithSeriesStyle(
            line ? new OfficeChartPointStyle(showOutline: false) : new OfficeChartPointStyle(noFill: true));
        content.Blocks.Add(new PowerPointLayoutChart(line ? OfficeChartType.Line : OfficeChartType.ColumnClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] { series })));
        var chart = Assert.Single(Assert.Single(PowerPointLayoutEngine.AddSlides(document, content)).Charts).GetChart();
        var properties = line ? chart.Descendants<C.LineChartSeries>().Single().ChartShapeProperties!
            : chart.Descendants<C.BarChartSeries>().Single().ChartShapeProperties!;
        Assert.NotNull(line ? properties.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>() : properties.GetFirstChild<A.NoFill>());
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FailedLayoutDoesNotLeaveAPartiallyBuiltSlideInTheSavedDeck(bool continuation) {
        using var document = PowerPointDocument.Create();
        var original = document.AddSlide();
        original.AddTextBox("Retained slide");
        var content = new PowerPointLayoutSlide { Title = "Unavailable illustration" };
        if (continuation) {
            content.Blocks.Add(new PowerPointLayoutList(new[] { "First completed content" }));
            content.Blocks.Add(new PowerPointLayoutList(new[] { "Second completed content" }));
        }
        content.Blocks.Add(new PowerPointLayoutImage(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png")));
        Assert.ThrowsAny<IOException>(() => PowerPointLayoutEngine.AddSlides(document, content));
        Assert.Same(original, Assert.Single(document.Slides));
        Assert.Empty(document.LayoutDeck.Diagnostics);
        using var saved = new MemoryStream();
        document.Save(saved);
        saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        Assert.Single(reopened.Slides);
        Assert.Empty(reopened.ValidateDocument());
    }
}
