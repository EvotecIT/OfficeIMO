using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(1, 0x80)]
    [InlineData(4, 0xF0)]
    [InlineData(8, 0x01)]
    public void CompleteContentValidationRejectsIndexedIconPixelsOutsideTheDeclaredPalette(
        int bitsPerPixel,
        int invalidPackedPixel) {
        byte[] valid = CreateIcon(CreateIndexedIconDib(bitsPerPixel, packedPixel: 0));
        byte[] invalid = CreateIcon(CreateIndexedIconDib(bitsPerPixel, (byte)invalidPackedPixel));

        Assert.True(OfficeImageReader.TryValidateContent(valid, "indexed.ico", out _));
        Assert.False(OfficeImageReader.TryValidateContent(invalid, "invalid-indexed.ico", out _));
    }

    [Fact]
    public void ExtendedVp8lUsesDecodedAlphaInsteadOfTheCodecHint() {
        byte[] simple = OfficeWebpCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var bytes = new List<byte> {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 0, 0, 0, 0,
            (byte)'W', (byte)'E', (byte)'B', (byte)'P'
        };
        bytes.AddRange(CreateWebpChunk("VP8X", new byte[10]));
        bytes.AddRange(simple.Skip(12));
        byte[] conservativeHint = bytes.ToArray();
        WriteInt32LittleEndian(conservativeHint, 4, conservativeHint.Length - 8);
        int vp8lOffset = FindWebpChunk(conservativeHint, "VP8L");
        conservativeHint[vp8lOffset + 12] |= 0x10;

        Assert.True(OfficeImageReader.TryIdentifyByContent(conservativeHint, "hint.webp", out _));
        Assert.True(OfficeImageReader.TryValidateContent(conservativeHint, "hint.webp", out _));

        byte[] incorrectContainerFlag = (byte[])conservativeHint.Clone();
        int vp8xOffset = FindWebpChunk(incorrectContainerFlag, "VP8X");
        incorrectContainerFlag[vp8xOffset + 8] |= 0x10;
        Assert.False(OfficeImageReader.TryIdentifyByContent(incorrectContainerFlag, "wrong-alpha-flag.webp", out _));
    }

    private static byte[] CreateIndexedIconDib(int bitsPerPixel, byte packedPixel) {
        const int headerSize = 40;
        const int paletteBytes = 4;
        const int xorBytes = 4;
        const int maskBytes = 4;
        var dib = new byte[headerSize + paletteBytes + xorBytes + maskBytes];
        WriteInt32LittleEndian(dib, 0, headerSize);
        WriteInt32LittleEndian(dib, 4, 1);
        WriteInt32LittleEndian(dib, 8, 2);
        WriteUInt16LittleEndian(dib, 12, 1);
        WriteUInt16LittleEndian(dib, 14, (ushort)bitsPerPixel);
        WriteInt32LittleEndian(dib, 20, xorBytes);
        WriteInt32LittleEndian(dib, 32, 1);
        dib[headerSize + paletteBytes] = packedPixel;
        return dib;
    }
}
