using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartPointStylesTests {
    [Fact]
    public void RadarHonorsIndependentGridlineAndCategoryLabelVisibility() {
        OfficeColor categoryGrid = OfficeColor.Parse("#B900D0");
        OfficeColor valueGrid = OfficeColor.Parse("#0077AA");
        var series = new OfficeChartSeries("Values", new[] { 1d, 2d, 3d }, null, null, null,
            showMarkers: false, connectLine: false);
        OfficeChartSnapshot Snapshot(bool category, bool value, bool labels) => new("", null,
            OfficeChartKind.Radar, new OfficeChartData(new[] { "A", "B", "C" }, new[] { series }), 320, 240,
            style: new OfficeChartStyle(categoryGridLineColor: categoryGrid, valueGridLineColor: valueGrid,
                showCategoryGridLines: category, showValueGridLines: value),
            layout: new OfficeChartLayout(showLegend: false, showCategoryAxisLabels: labels));
        OfficeDrawing hidden = OfficeChartDrawingRenderer.Render(Snapshot(false, false, false));
        Assert.DoesNotContain(hidden.Shapes, shape => shape.Shape.StrokeColor == categoryGrid || shape.Shape.StrokeColor == valueGrid);
        Assert.DoesNotContain(hidden.Elements.OfType<OfficeDrawingText>(), item => item.Text is "A" or "B" or "C");
        OfficeDrawing visible = OfficeChartDrawingRenderer.Render(Snapshot(true, true, true));
        Assert.Contains(visible.Shapes, shape => shape.Shape.StrokeColor == categoryGrid);
        Assert.Contains(visible.Shapes, shape => shape.Shape.StrokeColor == valueGrid);
        Assert.Contains(visible.Elements.OfType<OfficeDrawingText>(), item => item.Text == "A");
    }

    [Fact]
    public void DoughnutLegendUsesTheFirstRenderableRingWhenOnlyInnerRingHasPointStyles() {
        var outer = new OfficeChartSeries("Outer", new[] { 3d, 2d }, null, OfficeColor.Parse("#1845A3"));
        var inner = new OfficeChartSeries("Inner", new[] { 4d, 1d })
            .WithPointStyles(new OfficeChartPointStyle?[] {
                new(fillColor: OfficeColor.Parse("#D02020")), null
            });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Doughnut, new OfficeChartData(new[] { "A", "B" }, new[] { outer, inner }),
            480, 300, layout: new OfficeChartLayout(showLegend: true)));
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("#1845A3", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#D02020", svg, StringComparison.OrdinalIgnoreCase);
        var legend = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "A");
        Assert.Contains(drawing.Shapes, shape => shape.Y <= legend.Y + legend.Height &&
            shape.Y + shape.Shape.Height >= legend.Y && shape.Shape.FillColor == OfficeColor.Parse("#1845A3"));
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie, false)]
    [InlineData(OfficeChartKind.Pie, true)]
    [InlineData(OfficeChartKind.Doughnut, false)]
    [InlineData(OfficeChartKind.Doughnut, true)]
    public void PointStyles_UnfilledSliceLabelsContrastWithTheVisibleBackground(OfficeChartKind kind, bool dark) {
        OfficeColor background = dark ? OfficeColor.Black : OfficeColor.White;
        var series = new OfficeChartSeries("Results", new[] { 7d, 0d }).WithPointStyles(
            new OfficeChartPointStyle?[] { new(noFill: true), null });
        OfficeDrawing Draw(OfficeColor? labelColor = null) => OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot("Results", null, kind,
                new OfficeChartData(new[] { "A", "B" }, new[] { series }), 640, 360,
                style: new OfficeChartStyle(backgroundColor: background, dataLabelTextColor: labelColor),
                layout: new OfficeChartLayout(showLegend: false, showDataLabels: true, showDataLabelValues: true)));
        var labels = Draw().Elements.OfType<OfficeDrawingText>().Where(text => text.Text == "7" || text.Text == "0").ToArray();
        Assert.Equal(2, labels.Length);
        Assert.All(labels, label => Assert.Equal(dark ? OfficeColor.White : OfficeColor.Black, label.Color));
        Assert.All(Draw(OfficeColor.Red).Elements.OfType<OfficeDrawingText>().Where(text => text.Text == "7" || text.Text == "0"),
            label => Assert.Equal(OfficeColor.Red, label.Color));
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void PointStyles_UnfilledRadialLabelsIgnoreTheUnpaintedPlotSurface(OfficeChartKind kind) {
        var series = new OfficeChartSeries("Results", new[] { 7d, 0d }).WithPointStyles(
            new OfficeChartPointStyle?[] { new(noFill: true), null });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Results", null,
            kind, new OfficeChartData(new[] { "A", "B" }, new[] { series }), 640, 360,
            style: new OfficeChartStyle(backgroundColor: OfficeColor.White,
                plotAreaBackgroundColor: OfficeColor.Black),
            layout: new OfficeChartLayout(showLegend: false, showDataLabels: true,
                showDataLabelValues: true)));
        Assert.Equal(OfficeColor.Black, Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(),
            label => label.Text == "7").Color);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PointStyles_WidthOnlyOutlinesRemainVisibleUnlessExplicitlyHidden(bool hidden) {
        var series = new OfficeChartSeries("Results", new[] { 8d }).WithPointStyles(
            new OfficeChartPointStyle?[] { new(outlineWidth: 2, showOutline: hidden ? false : null) });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Results", null,
            OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] { series }), 320, 240,
            layout: new OfficeChartLayout(showLegend: false)));
        var bar = Assert.Single(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Rectangle &&
            shape.Shape.FillColor == OfficeChartDrawingRenderer.GetSeriesColor(0) && shape.Shape.Height > 50);
        Assert.Equal(hidden ? null : OfficeColor.Black, bar.Shape.StrokeColor);
        Assert.Equal(hidden ? 0 : 2, bar.Shape.StrokeWidth);
    }

    [Fact]
    public void PointStyles_CopyKeepsBubbleAndSeriesContracts() {
        var source = OfficeChartSeries.CreateBubble("Results", new[] { 1d }, new[] { 2d }, new[] { 3d },
            showInLegend: false, markerOutlineColor: OfficeColor.Black, markerOutlineWidth: 2,
            showMarkerOutline: false);
        var style = new OfficeChartPointStyle(noFill: true, outlineColor: OfficeColor.Black);
        var input = new OfficeChartPointStyle?[] { style };
        OfficeChartSeries copy = source.WithPointStyles(input);
        input[0] = null;
        Assert.Same(style, copy.PointStyles![0]);
        Assert.Null(source.PointStyles);
        Assert.Equal(source.BubbleSizes, copy.BubbleSizes);
        Assert.Equal(source.XValues, copy.XValues);
        Assert.Equal(source.Values, copy.Values);
        Assert.Equal(source.RenderKind, copy.RenderKind);
        Assert.False(copy.ShowInLegend);
        Assert.False(copy.ShowMarkerOutline);
        Assert.Equal(source.MarkerOutlineWidth, copy.MarkerOutlineWidth);
        Assert.Null(copy.WithPointStyles(null).PointStyles);
        Assert.Throws<ArgumentException>(() => source.WithPointStyles(Array.Empty<OfficeChartPointStyle?>()));
    }

    [Theory]
    [InlineData(OfficeChartHatchPattern.Horizontal)]
    [InlineData(OfficeChartHatchPattern.Vertical)]
    [InlineData(OfficeChartHatchPattern.ForwardDiagonal)]
    [InlineData(OfficeChartHatchPattern.BackwardDiagonal)]
    [InlineData(OfficeChartHatchPattern.Cross)]
    [InlineData(OfficeChartHatchPattern.DiagonalCross)]
    public void PointStyles_HatchesCoverEveryQuadrantOfThePoint(OfficeChartHatchPattern hatch) {
        OfficeColor ink = OfficeColor.Parse("#7300A3");
        var series = new OfficeChartSeries("Results", new[] { 8d }).WithPointStyles(
            new OfficeChartPointStyle?[] { new(hatch: hatch, hatchColor: ink) });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Results", null,
            OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] { series }), 320, 240,
            style: new OfficeChartStyle(backgroundColor: OfficeColor.Parse("#EEEEEE")),
            layout: new OfficeChartLayout(showLegend: false)));
        var bar = drawing.Shapes.Single(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle &&
            shape.Shape.FillColor == OfficeColor.White && shape.Shape.Height > 50);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++) {
                int pixels = 0;
                int left = (int)(bar.X + column * bar.Shape.Width / 2) + 3;
                int top = (int)(bar.Y + row * bar.Shape.Height / 2) + 3;
                int right = (int)(bar.X + (column + 1) * bar.Shape.Width / 2) - 3;
                int bottom = (int)(bar.Y + (row + 1) * bar.Shape.Height / 2) - 3;
                for (int y = top; y < bottom; y++)
                    for (int x = left; x < right; x++)
                        if (IsPurpleStroke(raster.GetPixel(x, y))) pixels++;
                Assert.True(pixels > 10, "Expected hatch coverage in quadrant " + row + "," + column);
            }
    }

    [Theory]
    [InlineData(OfficeChartKind.BarClustered, OfficeChartHatchPattern.Horizontal)]
    [InlineData(OfficeChartKind.ColumnClustered, OfficeChartHatchPattern.Vertical)]
    public void PointStyles_ThinBarsAndColumnsKeepTheirHatchDirection(
        OfficeChartKind kind, OfficeChartHatchPattern hatch) {
        OfficeColor ink = OfficeColor.Parse("#B900D0");
        string[] categories = Enumerable.Range(0, 60).Select(index => index.ToString()).ToArray();
        var styles = new OfficeChartPointStyle?[categories.Length];
        styles[categories.Length / 2] = new OfficeChartPointStyle(hatch: hatch, hatchColor: ink);
        var series = new OfficeChartSeries("Results", Enumerable.Repeat(8d, categories.Length).ToArray())
            .WithPointStyles(styles);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            kind, new OfficeChartData(categories, new[] { series }), 320, 240,
            layout: new OfficeChartLayout(showLegend: false)));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        int painted = 0;
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++)
            {
                var pixel = raster.GetPixel(x, y);
                if (pixel.R > pixel.G + 70 && pixel.B > pixel.G + 70 && pixel.R > 180) painted++;
            }
        Assert.True(painted > 0, "The requested hatch direction must remain visible on a thin point. SVG has ink: " +
            OfficeDrawingSvgExporter.ToSvg(drawing).Contains("#B900D0", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(OfficeChartMarkerShape.Dash)]
    [InlineData(OfficeChartMarkerShape.Plus)]
    [InlineData(OfficeChartMarkerShape.X)]
    public void PointStyles_LineOnlyScatterMarkersUseHatchForeground(OfficeChartMarkerShape markerShape) {
        OfficeColor ink = OfficeColor.Parse("#B900D0");
        var series = new OfficeChartSeries("Results", new[] { 4d }, new[] { 2d }, null, null,
            showMarkers: true, connectLine: false, markerShape: markerShape)
            .WithPointStyles(new OfficeChartPointStyle?[] {
                new(hatch: OfficeChartHatchPattern.Cross, hatchColor: ink)
            });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Scatter, new OfficeChartData(new[] { "2" }, new[] { series }), 320, 240,
            layout: new OfficeChartLayout(showLegend: false)));
        Assert.Contains(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == ink);
    }

    [Theory]
    [InlineData(OfficeChartMarkerShape.Dash)]
    [InlineData(OfficeChartMarkerShape.Plus)]
    [InlineData(OfficeChartMarkerShape.X)]
    public void PointStyles_LineOnlyScatterMarkersUsePointFillWhenNoOutlineIsSpecified(OfficeChartMarkerShape markerShape) {
        OfficeColor ink = OfficeColor.Parse("#B900D0");
        var series = new OfficeChartSeries("Results", new[] { 4d }, new[] { 2d }, null, null,
            showMarkers: true, connectLine: false, markerShape: markerShape)
            .WithPointStyles(new OfficeChartPointStyle?[] { new(fillColor: ink) });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Scatter, new OfficeChartData(new[] { "2" }, new[] { series }), 320, 240,
            layout: new OfficeChartLayout(showLegend: false)));
        Assert.Contains(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == ink);
    }

    [Fact]
    public void PointStyles_EnableBubbleOutlineOverDisabledSeriesOutline() {
        var series = OfficeChartSeries.CreateBubble("Results", new[] { 1d }, new[] { 2d }, new[] { 3d },
            showMarkerOutline: false).WithPointStyles(new OfficeChartPointStyle?[] { new(noFill: true, showOutline: true) });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Results", null,
            OfficeChartKind.Bubble, new OfficeChartData(new[] { "A" }, new[] { series }), 320, 240,
            layout: new OfficeChartLayout(showLegend: false)));
        var bubble = Assert.Single(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Ellipse);
        Assert.Null(bubble.Shape.FillColor);
        Assert.Equal(OfficeColor.Black, bubble.Shape.StrokeColor);
        Assert.True(bubble.Shape.StrokeWidth > 0);
    }

    [Fact]
    public void PointStyles_RejectContradictoryAndUnboundedAppearance() {
        Assert.Throws<ArgumentException>(() => new OfficeChartPointStyle(OfficeColor.Black, noFill: true));
        Assert.Throws<ArgumentException>(() => new OfficeChartPointStyle(hatch: OfficeChartHatchPattern.Cross));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartPointStyle(outlineWidth: double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeChartPointStyle(outlineWidth: 1585));
    }

    [Theory]
    [InlineData(OfficeChartKind.Area)]
    [InlineData(OfficeChartKind.AreaStacked)]
    [InlineData(OfficeChartKind.AreaStacked100)]
    public void PointStyles_ReportUnsupportedAreaAppearanceWithoutOmittingTheChart(OfficeChartKind kind) {
        var series = new OfficeChartSeries("Results", new[] { 3d, 2d, 1d }).WithPointStyles(
            new OfficeChartPointStyle?[] { null, new(noFill: true), null });
        var snapshot = new OfficeChartSnapshot("Results", null, kind,
            new OfficeChartData(new[] { "A", "B", "C" }, new[] { series }), 640, 360);
        OfficeChartRenderingResult result = OfficeChartDrawingRenderer.RenderWithQuality(snapshot);
        Assert.Contains(result.QualityReport.Issues, issue => issue.Kind == OfficeDrawingQualityIssueKind.UnsupportedAppearance);
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot, false, diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "ChartPointStylesUnsupported");
        Assert.NotEmpty(drawing.Shapes);
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void PointStyles_StyleHiddenCategoryLegendByOriginalPointIndex(OfficeChartKind kind) {
        var hatch = new OfficeChartPointStyle(OfficeColor.White, hatch: OfficeChartHatchPattern.Cross,
            hatchColor: OfficeColor.Parse("#7300A3"), outlineColor: OfficeColor.Black);
        var series = new OfficeChartSeries("Results", new[] { 3d, 2d, 1d })
            .WithPointStyles(new OfficeChartPointStyle?[] { null, null, hatch });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Results", null, kind,
            new OfficeChartData(new[] { "A", "B", "C" }, new[] { series }), 640, 360,
            layout: new OfficeChartLayout(hiddenCategoryLegendIndexes: new[] { 1 })));
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("#7300A3", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("clipPath", svg, StringComparison.Ordinal);
        Assert.DoesNotContain(">B<", svg, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void PointStyles_UnfilledCategoryKeepsItsInheritedLegendOutline(OfficeChartKind kind) {
        var series = new OfficeChartSeries("Results", new[] { 8D })
            .WithPointStyles(new OfficeChartPointStyle?[] { new(noFill: true) });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Results", null, kind,
            new OfficeChartData(new[] { "Could not evaluate" }, new[] { series }), 640D, 360D,
            style: new OfficeChartStyle(backgroundColor: OfficeColor.Black),
            layout: new OfficeChartLayout(showLegend: true)));
        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.FillColor == null &&
            shape.Shape.StrokeColor == OfficeColor.White &&
            shape.Shape.Width <= 20D && shape.Shape.Height <= 20D);
    }
    // Thin hatch strokes blend with their background during antialiasing.
    private static bool IsPurpleStroke(OfficeColor pixel) =>
        pixel.R < 200 && pixel.G < 150 && pixel.B > pixel.G + 40 && pixel.R > pixel.G + 20;
}
