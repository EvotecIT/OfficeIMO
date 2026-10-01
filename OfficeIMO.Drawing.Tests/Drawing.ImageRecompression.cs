using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingImageRecompressionTests {
    [Theory]
    [InlineData(OfficeImageOptimizationMode.Downsample, 32, 16)]
    [InlineData(OfficeImageOptimizationMode.Recompress, 128, 64)]
    [InlineData(OfficeImageOptimizationMode.DownsampleAndRecompress, 32, 16)]
    public void ModesControlPixelReductionAndJpegEncoding(OfficeImageOptimizationMode mode, int width, int height) {
        byte[] source = CreateJpeg();
        var result = OfficeImageOptimizer.Optimize(source, new OfficeImageOptimizationRequest(32, 16) {
            Mode = mode, JpegQuality = 35, KeepOriginalWhenNotSmaller = false
        });
        Assert.True(result.Changed);
        Assert.Equal(width, result.Final.Width);
        Assert.Equal(height, result.Final.Height);
        Assert.True(result.Bytes.Length < source.Length);
        Assert.True(OfficeImageReader.TryValidateContent(result.Bytes, "optimized.jpg", out _));
    }

    [Fact]
    public void RecompressionIsExplicitAtUnchangedDimensions() {
        byte[] source = CreateJpeg();
        var unchanged = OfficeImageOptimizer.Optimize(source, new OfficeImageOptimizationRequest(128, 64) {
            JpegQuality = 35
        });
        var compressed = OfficeImageOptimizer.Optimize(source, new OfficeImageOptimizationRequest(128, 64) {
            Mode = OfficeImageOptimizationMode.Recompress, JpegQuality = 35
        });
        Assert.Equal(OfficeImageOptimizationStatus.AlreadySuitable, unchanged.Status);
        Assert.Equal(source, unchanged.Bytes);
        Assert.True(compressed.Changed);
        Assert.Equal(128, compressed.Final.Width);
        Assert.Equal(64, compressed.Final.Height);
        Assert.True(compressed.BytesSaved > 0);
    }

    private static byte[] CreateJpeg() {
        var image = new OfficeRasterImage(128, 64);
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++)
                image.SetPixel(x, y, OfficeColor.FromRgb((byte)(x * 17), (byte)(y * 31), (byte)(x * y)));
        return OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions {
            Quality = 98, Subsampling = OfficeJpegSubsampling.Y444
        });
    }
}
