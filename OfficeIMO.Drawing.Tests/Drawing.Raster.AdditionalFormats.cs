using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class AdditionalRasterFormatTests {
    [Fact]
    public void BitfieldBitmapWithoutAlphaMaskTreatsFourthByteAsPadding() {
        byte[] bmp = new byte[82]; bmp[0] = 66; bmp[1] = 77; Put(bmp, 2, 82, 4); Put(bmp, 10, 74, 4);
        Put(bmp, 14, 56, 4); Put(bmp, 18, 2, 4); Put(bmp, 22, 1, 4); Put(bmp, 26, 1, 2); Put(bmp, 28, 32, 2); Put(bmp, 30, 3, 4); Put(bmp, 34, 8, 4);
        Put(bmp, 54, 0xFF0000, 4); Put(bmp, 58, 0xFF00, 4); Put(bmp, 62, 0xFF, 4);
        bmp[74] = 20; bmp[75] = 30; bmp[76] = 40; bmp[77] = 17; bmp[78] = 50; bmp[79] = 60; bmp[80] = 70; bmp[81] = 0;
        Assert.True(OfficeRasterImageDecoder.TryDecode(bmp, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.FromRgb(40, 30, 20), image!.GetPixel(0, 0)); Assert.Equal(OfficeColor.FromRgb(70, 60, 50), image.GetPixel(1, 0));
        Put(bmp, 66, 0xFF000000, 4);
        Assert.True(OfficeRasterImageDecoder.TryDecode(bmp, out OfficeRasterImage? alpha));
        Assert.Equal((byte)17, alpha!.GetPixel(0, 0).A); Assert.Equal((byte)0, alpha.GetPixel(1, 0).A);
    }
    [Fact]
    public void PortableBitmapAcceptsCompactPlainPixelsAndCommentsAtRawHeaderBoundary() {
        byte[] plain = System.Text.Encoding.ASCII.GetBytes("P1\n2 1\n01\n");
        Assert.True(OfficeRasterImageDecoder.TryDecode(plain, out OfficeRasterImage? compact));
        Assert.Equal(OfficeColor.White, compact!.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Black, compact.GetPixel(1, 0));
        foreach (string header in new[] { "P4\n1 1#comment\n", "P5\n1 1\n255#comment\n", "P6\n1 1\n255#comment\n" }) {
            byte[] prefix = System.Text.Encoding.ASCII.GetBytes(header);
            byte[] payload = header[1] == '4' ? new byte[] { 0x80 } : header[1] == '5' ? new byte[] { 127 } : new byte[] { 10, 20, 30 };
            var encoded = new byte[prefix.Length + payload.Length];
            Buffer.BlockCopy(prefix, 0, encoded, 0, prefix.Length);
            Buffer.BlockCopy(payload, 0, encoded, prefix.Length, payload.Length);
            Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded));
            OfficeColor expected = header[1] == '4' ? OfficeColor.Black : header[1] == '5' ? OfficeColor.FromRgb(127, 127, 127) : OfficeColor.FromRgb(10, 20, 30);
            Assert.Equal(expected, decoded!.GetPixel(0, 0));
        }
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Bmp)]
    [InlineData(OfficeImageExportFormat.Tga)]
    [InlineData(OfficeImageExportFormat.Icon)]
    public void LosslessFormatsPreserveStraightAlphaIncludingFullyTransparentImages(OfficeImageExportFormat format) {
        var image = new OfficeRasterImage(3, 2, OfficeColor.FromRgba(20, 40, 80, 0)); image.SetPixel(1, 0, OfficeColor.FromRgba(250, 120, 30, 77));
        byte[] encoded = OfficeRasterImageEncoder.Encode(image, format);
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded)); Assert.Equal(image.GetPixels(), decoded!.GetPixels());
        var transparent = new OfficeRasterImage(2, 1, OfficeColor.FromRgba(17, 33, 66, 0));
        Assert.True(OfficeRasterImageDecoder.TryDecode(OfficeRasterImageEncoder.Encode(transparent, format), out OfficeRasterImage? empty)); Assert.Equal(transparent.GetPixels(), empty!.GetPixels());
    }
    [Fact]
    public void MultiResolutionIconRetainsSelectedDimensionsAndLegacyMask() {
        byte[] icon = OfficeIconEncoder.Encode(new[] { new OfficeRasterImage(16, 16, OfficeColor.Red), new OfficeRasterImage(32, 32, OfficeColor.Blue) });
        Assert.True(OfficeRasterContainerInspector.TryInspect(icon, out OfficeRasterContainerInfo? container)); Assert.Equal(2, container!.Count);
        Assert.True(OfficeRasterImageDecoder.TryDecode(icon, new OfficeRasterDecodeOptions { FrameIndex = 1 }, out OfficeRasterImage? large, out _)); Assert.Equal(32, large!.Width); Assert.Equal(OfficeColor.Blue, large.GetPixel(0, 0));
        // An independently assembled four-bit DIB icon: red, green, and a transparency mask over the green pixel.
        byte[] dib = new byte[78]; dib[2] = 1; dib[4] = 1; dib[6] = 2; dib[7] = 1; Put(dib, 10, 1, 2); Put(dib, 12, 4, 2); Put(dib, 14, 56, 4); Put(dib, 18, 22, 4);
        Put(dib, 22, 40, 4); Put(dib, 26, 2, 4); Put(dib, 30, 2, 4); Put(dib, 34, 1, 2); Put(dib, 36, 4, 2); Put(dib, 54, 2, 4); dib[64] = 255; dib[67] = 255; dib[70] = 1; dib[74] = 64;
        Assert.True(OfficeRasterImageDecoder.TryDecode(dib, out OfficeRasterImage? legacy)); Assert.Equal(OfficeColor.Red, legacy!.GetPixel(0, 0)); Assert.Equal((byte)0, legacy.GetPixel(1, 0).A);
    }
    [Fact]
    public void PortableMapsDecodeIndependentAsciiAndBinaryFixturesAndThresholdWhiteCompositedAlpha() {
        byte[] ascii = Encoding.ASCII.GetBytes("P3\n# independent text fixture\n2 1\n255\n255 0 0 0 255 0\n");
        Assert.True(OfficeRasterImageDecoder.TryDecode(ascii, out OfficeRasterImage? rgb)); Assert.Equal(OfficeColor.Red, rgb!.GetPixel(0, 0)); Assert.Equal(OfficeColor.FromRgb(0, 255, 0), rgb.GetPixel(1, 0));
        byte[] gray = Encoding.ASCII.GetBytes("P5\n2 1\n65535\n").Concat(new byte[] { 0, 0, 255, 255 }).ToArray(); Assert.True(OfficeRasterImageDecoder.TryDecode(gray, out OfficeRasterImage? grayscale)); Assert.Equal(OfficeColor.Black, grayscale!.GetPixel(0, 0)); Assert.Equal(OfficeColor.White, grayscale.GetPixel(1, 0));
        var source = new OfficeRasterImage(9, 1, OfficeColor.Black); source.SetPixel(0, 0, OfficeColor.FromRgba(0, 0, 0, 0));
        Assert.True(OfficeRasterImageDecoder.TryDecode(OfficeRasterImageEncoder.Encode(source, OfficeImageExportFormat.Pbm), out OfficeRasterImage? binary)); Assert.Equal(OfficeColor.White, binary!.GetPixel(0, 0)); Assert.Equal(OfficeColor.Black, binary.GetPixel(8, 0));
        Assert.False(OfficeRasterImageDecoder.TryDecode(ascii, new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1 }, out _, out _));
    }
    [Fact]
    public void TgaRlePacketsRespectOriginAndRejectRunsOutsideTheImage() {
        byte[] tga = new byte[22]; tga[2] = 10; tga[12] = 2; tga[14] = 1; tga[16] = 24; tga[17] = 32; tga[18] = 129; tga[21] = 255;
        Assert.True(OfficeRasterImageDecoder.TryDecode(tga, out OfficeRasterImage? red)); Assert.Equal(OfficeColor.Red, red!.GetPixel(0, 0)); Assert.Equal(OfficeColor.Red, red.GetPixel(1, 0));
        tga[18] = 130; Assert.False(OfficeRasterImageDecoder.TryDecode(tga, out _));
    }
    [Fact]
    public void FastPngCompressionAndEncodedByteLimitsAreRealWriterControls() {
        var image = new OfficeRasterImage(5, 4, OfficeColor.Blue); var options = new OfficeRasterEncodingOptions { Png = new OfficePngEncodeOptions { Compression = OfficePngCompression.Fastest } };
        byte[] encoded = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png, options); Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded)); Assert.Equal(image.GetPixels(), decoded!.GetPixels());
        Assert.ThrowsAny<Exception>(() => OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Bmp, null, 20));
    }
    private static void Put(byte[] bytes, int offset, uint value, int size) { for (int i = 0; i < size; i++) bytes[offset + i] = (byte)(value >> (i * 8)); }
}
