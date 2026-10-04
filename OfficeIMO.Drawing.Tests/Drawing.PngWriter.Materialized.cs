using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingPngWriterTests {
    [Theory]
    [InlineData(512, 512, 0)]
    [InlineData(512, 512, 1)]
    [InlineData(1024, 256, 2)]
    [InlineData(2048, 768, 2)]
    [InlineData(2048, 1024, 3)]
    [InlineData(1024, 256, 4)]
    [InlineData(2048, 1536, 4)]
    public void MaterializedOptimalPngMatchesStreamingForStructuredAndLargeRgba(int width, int height, int pattern) {
        byte[] pixels = CreateMaterializedPixels(width, height, pattern);
        var image = OfficeRasterImage.FromRgba32(width, height, pixels);
        var options = new OfficePngEncodeOptions { DpiX = 144, DpiY = 120 };
        using var reference = new MemoryStream();
        OfficePngWriter.EncodeTo(image, reference, options);

        byte[] actual = OfficePngWriter.EncodeRgba(width, height, pixels, options);

        // The independent streaming path protects minimum-size selection, full IDAT
        // framing and replacement of a longer candidate without retaining its tail.
        Assert.Equal(reference.ToArray(), actual);
        Assert.True(OfficePngReader.TryDecode(actual, out OfficeRasterImage? decoded));
        Assert.Equal(pixels, decoded!.GetPixels());
        OfficeImageInfo info = OfficeImageReader.Identify(actual);
        Assert.InRange(info.DpiX, 143.98, 144.02);
        Assert.InRange(info.DpiY, 119.98, 120.02);
    }

    [Theory]
    [InlineData(OfficePngCompression.Optimal)]
    [InlineData(OfficePngCompression.Stored)]
    public void MaterializedPngEntryPointsMatchStreamingWithAndWithoutResolution(OfficePngCompression compression) {
        const int width = 512, height = 512;
        byte[] pixels = CreateMaterializedPixels(width, height, pattern: 0);
        var image = OfficeRasterImage.FromRgba32(width, height, pixels);
        var options = new OfficePngEncodeOptions { Compression = compression, DpiX = 144, DpiY = 120 };
        using var plain = new MemoryStream();
        using var physical = new MemoryStream();
        OfficePngWriter.EncodeTo(image, plain, compression);
        OfficePngWriter.EncodeTo(image, physical, options);

        byte[] expected = plain.ToArray();
        Assert.Equal(expected, OfficePngWriter.Encode(image, compression));
        Assert.Equal(expected, OfficePngWriter.Encode(image, CancellationToken.None, compression));
        Assert.Equal(expected, OfficePngWriter.EncodeRgba(width, height, pixels, compression));
        Assert.Equal(physical.ToArray(), OfficePngWriter.Encode(image, options));
        Assert.Equal(physical.ToArray(), OfficePngWriter.EncodeRgba(width, height, pixels, options));
        options.WritePhysicalResolution = false;
        Assert.Equal(expected, OfficePngWriter.Encode(image, options));
        Assert.Equal(expected, OfficePngWriter.EncodeRgba(width, height, pixels, options));
    }

    private static byte[] CreateMaterializedPixels(int width, int height, int pattern) {
        byte[] pixels = new byte[width * height * 4];
        uint state = 0x915DF32B;
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                int offset = (y * width + x) * 4;
                for (int channel = 0; channel < 4; channel++) {
                    state ^= state << 13;
                    state ^= state >> 17;
                    state ^= state << 5;
                    byte value;
                    if (pattern == 0) {
                        value = channel == 3 ? (byte)255 : unchecked((byte)(x * 3 + y * 5 + channel * 31));
                    } else if (pattern == 1) {
                        value = channel == 3 ? (byte)255 :
                            (byte)(x % 19 < 10 && y % 13 < 7 ? 17 * ((x + y) % 16) : 255);
                    } else if (pattern == 4) {
                        value = channel == 3 ? (byte)255 : (byte)((state & 15) * 17);
                    } else if (pattern == 3 && y > 0) {
                        value = unchecked((byte)(pixels[offset - width * 4 + channel] + (state & 31) - 16));
                    } else {
                        value = (byte)state;
                    }
                    pixels[offset + channel] = value;
                }
            }
        }
        return pixels;
    }
}
