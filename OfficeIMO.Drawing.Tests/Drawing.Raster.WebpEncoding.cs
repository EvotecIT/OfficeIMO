using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpEncodingTests {
    [Fact]
    public void HalfPredictorTruncatesNegativeOddDifferencesTowardZero() {
        uint left = 0x10121416U, top = 0x10121416U, topLeft = 0x1317191DU;
        uint expected = 0;
        for (int shift = 0; shift < 32; shift += 8) {
            int average = ((int)(left >> shift) & 255);
            int delta = average - ((int)(topLeft >> shift) & 255);
            int half = delta < 0 ? -((-delta) / 2) : delta / 2;
            int channel = System.Math.Max(0, System.Math.Min(255, average + half));
            expected |= (uint)channel << shift;
        }
        Assert.Equal(expected, OfficeWebpCodec.PredictVp8l(13, left, top, topLeft, 0));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void FrequencyCodedWebpPreservesSparseChannelsAndHiddenTransparentColor(int variant) {
        var image = new OfficeRasterImage(67, 35);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                byte red = variant == 0 ? (byte)33 : (byte)(x * 17 + y * 11);
                byte green = variant == 1 ? (byte)0 : (byte)(x * 3 + y * 7);
                byte blue = variant == 2 ? (byte)255 : (byte)(x * 13 + y * 19);
                byte alpha = variant == 3 ? (byte)0 : (byte)255;
                image.SetPixel(x, y, OfficeColor.FromRgba(red, green, blue, alpha));
            }
        }
        byte[] original = image.GetPixels();
        byte[] encoded = OfficeWebpCodec.Encode(image);
        Assert.True(OfficeWebpCodec.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(original, decoded!.GetPixels());
        Assert.Equal(original, image.GetPixels());
        Assert.Equal(encoded, OfficeWebpCodec.Encode(image));
    }
}
