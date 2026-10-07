using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingPngWriterTests {
    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(7)]
    [InlineData(15)]
    [InlineData(16)]
    [InlineData(17)]
    [InlineData(1023)]
    [InlineData(1024)]
    [InlineData(1025)]
    [InlineData(4095)]
    [InlineData(4096)]
    [InlineData(4097)]
    [InlineData(16385)]
    public void OpaqueRgbPngPreservesChannelsAcrossRowTailsAndEncodingSurfaces(int width) {
        var image = new OfficeRasterImage(width, 3);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < width; x++) {
                image.SetPixel(x, y, OfficeColor.FromRgb(
                    unchecked((byte)(x * 127 + y * 128 + 1)),
                    unchecked((byte)(x * 31 - y * 128 + 2)),
                    unchecked((byte)(x * 255 + y + 3))));
            }
        }
        byte[] pixels = image.GetPixels();
        var options = new OfficePngEncodeOptions { DpiX = 300, DpiY = 144 };
        byte[] png = OfficePngWriter.Encode(image, options);
        using var stream = new MemoryStream();
        OfficePngWriter.EncodeTo(image, stream, options);
        Assert.Equal(png, stream.ToArray());
        Assert.Equal(png, OfficePngWriter.EncodeRgba(width, image.Height, pixels, options));
#if NET8_0_OR_GREATER
        var writer = new ArrayBufferWriter<byte>();
        OfficePngWriter.EncodeTo(image, writer, options);
        Assert.Equal(png, writer.WrittenSpan.ToArray());
#endif
        Assert.Equal(8, png[24]);
        Assert.Equal(2, png[25]);
        Assert.True(OfficePngReader.TryDecode(png, out var decoded));
        Assert.Equal(pixels, decoded!.GetPixels());
        Assert.Equal(pixels, image.GetPixels());
        OfficeImageInfo info = OfficeImageReader.Identify(png);
        Assert.InRange(info.DpiX, 299.98, 300.02);
        Assert.InRange(info.DpiY, 143.98, 144.02);

        byte[] stored = OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        Assert.Equal(6, stored[25]);
        Assert.True(OfficePngReader.TryDecode(stored, out decoded));
        Assert.Equal(pixels, decoded!.GetPixels());
    }
}
