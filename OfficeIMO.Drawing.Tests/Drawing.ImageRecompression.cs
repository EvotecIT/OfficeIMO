using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingImageRecompressionTests {
    [Theory]
    [InlineData(1, 0, 1, 2, 3)]
    [InlineData(2, 1, 0, 3, 2)]
    [InlineData(3, 3, 2, 1, 0)]
    [InlineData(4, 2, 3, 0, 1)]
    [InlineData(5, 0, 2, 1, 3)]
    [InlineData(6, 2, 0, 3, 1)]
    [InlineData(7, 3, 1, 2, 0)]
    [InlineData(8, 1, 3, 0, 2)]
    public void RecompressionPreservesExifOrientedPixels(int orientation, int topLeft, int topRight,
        int bottomLeft, int bottomRight) {
        var colors = new[] { OfficeColor.Red, OfficeColor.Green, OfficeColor.Blue, OfficeColor.Yellow };
        var raster = new OfficeRasterImage(64, 32);
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                raster.SetPixel(x, y, colors[(y < 16 ? 0 : 2) + (x < 32 ? 0 : 1)]);
        byte[] exif = {
            (byte)'I', (byte)'I', 0x2A, 0, 8, 0, 0, 0, 1, 0,
            0x12, 1, 3, 0, 1, 0, 0, 0, (byte)orientation, 0, 0, 0, 0, 0, 0, 0
        };
        byte[] source = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
            Quality = 98, Subsampling = OfficeJpegSubsampling.Y444, Metadata = new OfficeJpegMetadata(exif: exif)
        });
        foreach (var mode in new[] { OfficeImageOptimizationMode.Recompress, OfficeImageOptimizationMode.DownsampleAndRecompress }) {
            var result = OfficeImageOptimizer.Optimize(source, new OfficeImageOptimizationRequest(64, 64) {
                Mode = mode, JpegQuality = 90, KeepOriginalWhenNotSmaller = false
            });
            Assert.True(result.Changed);
            Assert.True(OfficeImageOrientationNormalizer.TryRead(result.Bytes, out var outputOrientation));
            Assert.Equal(OfficeImageOrientation.Normal, outputOrientation);
            var decoded = OfficeJpegCodec.Decode(result.Bytes);
            Assert.Equal(orientation >= 5 ? 32 : 64, decoded.Width);
            Assert.Equal(orientation >= 5 ? 64 : 32, decoded.Height);
            AssertColor(colors[topLeft], decoded.GetPixel(8, 8));
            AssertColor(colors[topRight], decoded.GetPixel(decoded.Width - 9, 8));
            AssertColor(colors[bottomLeft], decoded.GetPixel(8, decoded.Height - 9));
            AssertColor(colors[bottomRight], decoded.GetPixel(decoded.Width - 9, decoded.Height - 9));
        }
    }

    private static void AssertColor(OfficeColor expected, OfficeColor actual) {
        Assert.InRange(Math.Abs(expected.R - actual.R), 0, 30);
        Assert.InRange(Math.Abs(expected.G - actual.G), 0, 30);
        Assert.InRange(Math.Abs(expected.B - actual.B), 0, 30);
    }

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
