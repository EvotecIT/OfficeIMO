using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegEncoderQualityTests {
    [Fact]
    public void DefaultHuffmanOptimizationChangesOnlyEntropyCoding() {
        var image = new OfficeRasterImage(127, 93);
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++)
            image.SetPixel(x, y, OfficeColor.FromRgb((byte)x, (byte)(x + y), (byte)(y * 2)));
        byte[] optimized = OfficeJpegCodec.Encode(image);
        byte[] standard = OfficeJpegCodec.Encode(image, new OfficeJpegEncodeOptions { OptimizeHuffman = false });
        Assert.True(optimized.Length < standard.Length);
        Assert.True(OfficeRasterImageDecoder.TryDecode(optimized, out var a));
        Assert.True(OfficeRasterImageDecoder.TryDecode(standard, out var b));
        Assert.Equal(a!.GetPixels(), b!.GetPixels());
        Assert.Equal(OfficeJpegSubsampling.Y444, new OfficeJpegEncodeOptions().Subsampling);
    }
}
