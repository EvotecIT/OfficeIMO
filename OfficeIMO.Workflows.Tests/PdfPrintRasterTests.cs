using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPrintRasterTests {
    [Fact]
    public void RowConversionMatchesTheSharedGrayscaleAndProducesGdiByteOrder() {
        OfficeColor[] colors = [OfficeColor.FromRgba(255, 0, 0, 255), OfficeColor.FromRgba(12, 200, 99, 128), OfficeColor.FromRgba(0, 0, 255, 0)];
        var raster = new OfficeRasterImage(3, 1);
        for (int x = 0; x < colors.Length; x++) raster.SetPixel(x, 0, colors[x]);
        byte[] bgra = raster.GetPixels();
        PdfPrintRaster.ApplyGrayscale(raster, CancellationToken.None);
        for (int x = 0; x < colors.Length; x++) Assert.Equal(OfficeColorTransforms.Grayscale(colors[x]), raster.GetPixel(x, 0));

        PdfPrintRaster.ConvertRows(bgra, 3, 1, grayscale: false, toBgra: true, CancellationToken.None);
        Assert.Equal(new byte[] { 0, 0, 255, 255, 99, 200, 12, 128, 255, 0, 0, 0 }, bgra);
    }

    [Theory]
    [InlineData(600, 300, 50_000_000, 300)] // Caller limit below the device.
    [InlineData(200, 300, 50_000_000, 200)] // Device below the caller limit.
    [InlineData(0, 300, 50_000_000, 300)]   // Unknown device resolution.
    [InlineData(600, 600, 8_700_000, 299)]  // A4 pixel budget.
    public void RasterResolutionHonoursDeviceCallerAndPixelLimits(double device, double maximum, long pixels, double expected) {
        Assert.Equal(expected, PdfPrintRaster.ResolveDpi(device, maximum, 595.28, 841.89, pixels));
    }
}
