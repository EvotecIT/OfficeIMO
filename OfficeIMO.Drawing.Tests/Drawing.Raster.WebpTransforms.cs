using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpTransformTests {
    [Fact]
    public void SelectedNeighborPreservesAllChannelDistancesAndTopOnTies() {
        Assert.Equal(0xFF000000U, OfficeWebpCodec.PredictVp8l(11, 0x00FFFFFFU, 0xFF000000U, 0x80808080U, 0));
        var random = new Random(11011);
        byte[] colors = new byte[12];
        for (int index = 0; index < 8192; index++) {
            random.NextBytes(colors);
            uint left = BitConverter.ToUInt32(colors, 0), top = BitConverter.ToUInt32(colors, 4), corner = BitConverter.ToUInt32(colors, 8);
            int leftDistance = 0, topDistance = 0;
            for (int shift = 0; shift < 32; shift += 8) {
                int l = (int)(left >> shift) & 255, t = (int)(top >> shift) & 255;
                int estimate = l + t - ((int)(corner >> shift) & 255);
                leftDistance += Math.Abs(estimate - l);
                topDistance += Math.Abs(estimate - t);
            }
            uint expected = leftDistance < topDistance ? left : top;
            Assert.Equal(expected, OfficeWebpCodec.PredictVp8l(11, left, top, corner, 0));
        }
        uint first = 0x11223344U, second = 0x44332211U;
        Assert.Equal(second, OfficeWebpCodec.PredictVp8l(11, first, second, 0x00000000U, 0));
    }

    [Theory]
    [InlineData(1, 39)]
    [InlineData(39, 1)]
    [InlineData(8, 9)]
    [InlineData(17, 13)]
    [InlineData(65, 33)]
    [InlineData(257, 131)]
    public void PredictorReconstructionPreservesRowsBordersAndHiddenTransparentColor(int width, int height) {
        var source = new OfficeRasterImage(width, height);
        var random = new Random(width * 997 + height);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                source.SetPixel(x, y, OfficeColor.FromRgba((byte)random.Next(256),
                    (byte)(x * 17 + y * 31), (byte)random.Next(256),
                    (byte)((x + y) % 3 == 0 ? 0 : random.Next(256))));
            }
        }
        byte[] expected = source.GetPixels();
        byte[] encoded = OfficeWebpCodec.Encode(source);
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out var decoded));
        Assert.NotNull(decoded);
        Assert.Equal(expected, decoded.GetPixels());
        Assert.Equal(expected, source.GetPixels());
    }

}
