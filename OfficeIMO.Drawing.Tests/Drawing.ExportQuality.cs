using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingExportQualityTests {
    [Fact]
    public void ImportedLargeChartTitleReservesItsFullTextHeight() {
        var snapshot = new OfficeChartSnapshot("Native chart", "Shared combination", OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1" }, new[] { new OfficeChartSeries("Volume", new[] { 100d }) }),
            480, 300, style: new OfficeChartStyle(titleFontSize: 24),
            layout: new OfficeChartLayout(showLegend: false));

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
        OfficeDrawingText title = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Shared combination");
        Assert.Equal(24, title.Font.Size);
        Assert.True(title.Height >= 30D);
    }

    [Fact]
    public void QualityAnalysisIncludesTransformedGroupTextAndOverflow() {
        var child = new OfficeDrawing(100, 100);
        child.AddText("Nested", 10, 10, 40, 20);
        child.AddShapeForClippedRendering(OfficeShape.Rectangle(30, 10), 90, 10);
        var drawing = new OfficeDrawing(50, 50);
        drawing.AddEffectDrawing(child, OfficeTransform.Scale(.5, .5));
        drawing.AddText("Outer", 5, 5, 20, 10);
        var issues = OfficeDrawingQualityAnalyzer.Analyze(drawing).Issues;
        Assert.Contains(issues, issue => issue.Kind == OfficeDrawingQualityIssueKind.TextOverlap && issue.ElementIndex == 0 && issue.RelatedElementIndex == 1);
        Assert.Contains(issues, issue => issue.Kind == OfficeDrawingQualityIssueKind.ElementOutsideBounds && issue.ElementIndex == 0);
    }

    [Fact]
    public void SmallChartQualityRenderingStillReportsOverlappingPointLabels() {
        var snapshot = new OfficeChartSnapshot("Small scatter", null, OfficeChartKind.Scatter,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 2d, 2d }, new[] { 1d, 1d }) }),
            90, 30, layout: new OfficeChartLayout(showDataLabels: true, showDataLabelValues: true));
        var result = OfficeChartDrawingRenderer.RenderWithQuality(snapshot, useMinimumCanvas: false);
        Assert.Contains(result.QualityReport.Issues, issue => issue.Kind == OfficeDrawingQualityIssueKind.TextOverlap);
    }

    [Fact]
    public void ChartQualityRenderingCanPreserveSmallAuthoredDimensions() {
        var snapshot = new OfficeChartSnapshot("Small chart", null, OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }), 90, 30);
        var authored = OfficeChartDrawingRenderer.RenderWithQuality(snapshot, useMinimumCanvas: false);
        Assert.Equal(90d, authored.Drawing.Width);
        Assert.Equal(30d, authored.Drawing.Height);
        var defaultCanvas = OfficeChartDrawingRenderer.RenderWithQuality(snapshot);
        Assert.Equal(240d, defaultCanvas.Drawing.Width);
        Assert.Equal(150d, defaultCanvas.Drawing.Height);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public void DrawingRasterExportHonorsTheEncodedByteLimit(OfficeImageExportFormat format) {
        var drawing = new OfficeDrawing(24, 16);
        var error = Assert.Throws<OfficeImageExportBatchLimitException>(() => drawing.ExportImage(format,
            new OfficeImageExportOptions { MaximumTotalEncodedBytes = 1 }));
        Assert.Equal(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), error.LimitName);
        var expected = drawing.ExportImage(format);
        var exact = drawing.ExportImage(format, new OfficeImageExportOptions { MaximumTotalEncodedBytes = expected.Bytes.LongLength });
        Assert.Equal(expected.Bytes, exact.Bytes);
    }

    [Theory]
    [InlineData(255)]
    [InlineData(128)]
    [InlineData(0)]
    public void SvgExportPaintsTheSelectedBackgroundBelowExistingContent(int alpha) {
        var drawing = new OfficeDrawing(24, 16);
        var foreground = OfficeShape.Rectangle(8, 8); foreground.FillColor = OfficeColor.Blue; foreground.StrokeWidth = 0;
        drawing.AddShape(foreground, 8, 4);
        var originalShape = drawing.Shapes.Single().Shape;
        var background = new OfficeColor(230, 120, 40, (byte)alpha);
        var options = new OfficeImageExportOptions { BackgroundColor = background };
        var result = drawing.ExportImage(OfficeImageExportFormat.Svg, options);
        Assert.True(OfficeSvgDrawingReader.TryRead(result.Bytes, out var imported));
        var raster = OfficeDrawingRasterRenderer.Render(imported!, new OfficeDrawingRasterRenderOptions { Background = OfficeColor.Transparent });
        Assert.Equal(alpha == 0 ? OfficeColor.Transparent : background, raster.GetPixel(2, 2));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(12, 8));
        Assert.Single(drawing.Elements); Assert.Same(originalShape, drawing.Shapes.Single().Shape);
    }

    [Fact]
    public void DefaultSvgBackgroundIsWhiteAndExplicitPageFillStillWins() {
        var drawing = new OfficeDrawing(24, 16);
        var empty = drawing.ExportImage(OfficeImageExportFormat.Svg);
        Assert.True(OfficeSvgDrawingReader.TryRead(empty.Bytes, out var imported));
        Assert.Equal(OfficeColor.White, OfficeDrawingRasterRenderer.Render(imported!,
            new OfficeDrawingRasterRenderOptions { Background = OfficeColor.Transparent }).GetPixel(2, 2));
        var page = OfficeShape.Rectangle(24, 16); page.FillColor = OfficeColor.Green; page.StrokeWidth = 0;
        drawing.AddShape(page, 0, 0);
        var filled = drawing.ExportImage(OfficeImageExportFormat.Svg, new OfficeImageExportOptions { BackgroundColor = OfficeColor.Red });
        Assert.True(OfficeSvgDrawingReader.TryRead(filled.Bytes, out imported));
        Assert.Equal(OfficeColor.Green, OfficeDrawingRasterRenderer.Render(imported!,
            new OfficeDrawingRasterRenderOptions { Background = OfficeColor.Transparent }).GetPixel(2, 2));
    }

    private sealed class PointOptions : OfficeImageExportOptions {
        public override double LogicalUnitsPerInch => 72;
    }

    [Theory]
    [InlineData(OfficeImageExportQuality.Preview, 96)]
    [InlineData(OfficeImageExportQuality.Screen, 192)]
    [InlineData(OfficeImageExportQuality.Print, 300)]
    public void DensityIsRenderedAndEncodedWithoutMutatingOptions(OfficeImageExportQuality quality, int dpi) {
        var drawing = new OfficeDrawing(72, 36);
        drawing.AddText("Quality", 2, 2, 68, 30, new OfficeFontInfo("Arial", 12), OfficeColor.Black);
        var options = new PointOptions().UseQuality(quality);
        var image = drawing.ExportImage(OfficeImageExportFormat.Png, options);
        Assert.Equal(dpi, image.Width); Assert.Equal(dpi / 2, image.Height);
        // PNG stores integer pixels per metre, with up to 0.0127 DPI rounding.
        Assert.InRange(image.DpiX, dpi - .013, dpi + .013);
        Assert.InRange(image.PhysicalWidthInches, .999, 1.001);
        Assert.Equal(1, options.Scale); Assert.Equal(dpi, options.TargetDpi);
        Assert.NotEmpty(image.Bytes);
    }

    [Fact]
    public void VectorExportRetainsTextAndPixelLimitsFailExplicitly() {
        var drawing = new OfficeDrawing(72, 36);
        drawing.AddText("Vector label", 0, 0, 72, 30, new OfficeFontInfo("Arial", 10), OfficeColor.Black);
        var options = new PointOptions().UseQuality(OfficeImageExportQuality.Print);
        var svg = drawing.ExportImage(OfficeImageExportFormat.Svg, options);
        Assert.Contains("Vector label", System.Text.Encoding.UTF8.GetString(svg.Bytes));
        options.MaximumRasterPixels = 100; options.RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw;
        Assert.Throws<OfficeImageExportLimitException>(() => drawing.ExportImage(OfficeImageExportFormat.Png, options));
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => drawing.ExportImage(OfficeImageExportFormat.Svg, cancellationToken: cancel.Token));
    }
}
