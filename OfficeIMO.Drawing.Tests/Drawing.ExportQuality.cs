using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingExportQualityTests {
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
