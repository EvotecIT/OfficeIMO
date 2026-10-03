using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PngRgbaScanlineTests {
    public static IEnumerable<object[]> FiltersAndWidths() {
        foreach (int filter in new[] { 0, 1, 2, 3, 4 }) {
            foreach (int width in new[] { 1, 16, 17, 1025 }) yield return new object[] { filter, width };
        }
    }

    [Theory]
    [MemberData(nameof(FiltersAndWidths))]
    public void EveryRgbaFilterPreservesFirstRowChannelsAndBlockTails(int filter, int width) {
        const int height = 3;
        int stride = width * 4;
        var expected = new byte[stride * height];
        new Random(1951 + width).NextBytes(expected);
        var scanlines = new byte[(stride + 1) * height];
        for (int row = 0; row < height; row++) {
            int offset = row * stride;
            int target = row * (stride + 1);
            scanlines[target++] = (byte)filter;
            for (int index = 0; index < stride; index++) {
                int left = index >= 4 ? expected[offset + index - 4] : 0;
                int above = row != 0 ? expected[offset - stride + index] : 0;
                int diagonal = row != 0 && index >= 4 ? expected[offset - stride + index - 4] : 0;
                int prediction = filter switch {
                    0 => 0,
                    1 => left,
                    2 => above,
                    3 => (left + above) / 2,
                    _ => Paeth(left, above, diagonal)
                };
                scanlines[target + index] = unchecked((byte)(expected[offset + index] - prediction));
            }
        }
        byte[] encoded = OfficePngWriter.EncodeScanlines(width, height, 8, 6, scanlines);

        Assert.True(OfficePngReader.TryDecode(encoded, out OfficeRasterImage? decoded));
        Assert.NotNull(decoded);
        Assert.Equal(expected, decoded!.GetPixels());
    }

    private static int Paeth(int left, int above, int diagonal) {
        int estimate = left + above - diagonal;
        int leftDistance = Math.Abs(estimate - left);
        int aboveDistance = Math.Abs(estimate - above);
        int diagonalDistance = Math.Abs(estimate - diagonal);
        return leftDistance <= aboveDistance && leftDistance <= diagonalDistance
            ? left
            : aboveDistance <= diagonalDistance ? above : diagonal;
    }
}
