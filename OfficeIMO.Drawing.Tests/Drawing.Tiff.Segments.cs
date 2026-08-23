using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTiffSegmentTests {
    [Fact]
    public void TiffDecoderDecodesPlanarRgbStrips() {
        byte[] tiff = CreatePlanarRgbTiff();

        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal(OfficeColor.Red, image!.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Lime, image.GetPixel(1, 0));
    }

    [Fact]
    public void TiffDecoderDecodesPaddedChunkyTiles() {
        byte[] tiff = CreateTiledRgbTiff();

        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal((3, 2), (image!.Width, image.Height));
        Assert.Equal(OfficeColor.Red, image.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Lime, image.GetPixel(1, 0));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(2, 0));
        Assert.Equal(OfficeColor.White, image.GetPixel(2, 1));
    }

    private static byte[] CreatePlanarRgbTiff() {
        const int entryCount = 10;
        const int ifdOffset = 8;
        int dataOffset = ifdOffset + 2 + entryCount * 12 + 4;
        int bitsOffset = dataOffset;
        int offsetsOffset = bitsOffset + 6;
        int countsOffset = offsetsOffset + 12;
        int pixelsOffset = countsOffset + 12;
        var output = new byte[pixelsOffset + 6];
        WriteHeader(output, ifdOffset, entryCount);
        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, 2);
        WriteEntry(output, ref entry, 257, 4, 1, 1);
        WriteEntry(output, ref entry, 258, 3, 3, bitsOffset);
        WriteEntry(output, ref entry, 259, 3, 1, 1);
        WriteEntry(output, ref entry, 262, 3, 1, 2);
        WriteEntry(output, ref entry, 273, 4, 3, offsetsOffset);
        WriteEntry(output, ref entry, 277, 3, 1, 3);
        WriteEntry(output, ref entry, 278, 4, 1, 1);
        WriteEntry(output, ref entry, 279, 4, 3, countsOffset);
        WriteEntry(output, ref entry, 284, 3, 1, 2);
        WriteUInt32(output, entry, 0);
        for (int index = 0; index < 3; index++) {
            WriteUInt16(output, bitsOffset + index * 2, 8);
            WriteUInt32(output, offsetsOffset + index * 4, pixelsOffset + index * 2);
            WriteUInt32(output, countsOffset + index * 4, 2);
        }
        output[pixelsOffset] = 255;
        output[pixelsOffset + 1] = 0;
        output[pixelsOffset + 2] = 0;
        output[pixelsOffset + 3] = 255;
        return output;
    }

    private static byte[] CreateTiledRgbTiff() {
        const int entryCount = 11;
        const int ifdOffset = 8;
        int dataOffset = ifdOffset + 2 + entryCount * 12 + 4;
        int bitsOffset = dataOffset;
        int offsetsOffset = bitsOffset + 6;
        int countsOffset = offsetsOffset + 8;
        int pixelsOffset = countsOffset + 8;
        var output = new byte[pixelsOffset + 24];
        WriteHeader(output, ifdOffset, entryCount);
        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, 3);
        WriteEntry(output, ref entry, 257, 4, 1, 2);
        WriteEntry(output, ref entry, 258, 3, 3, bitsOffset);
        WriteEntry(output, ref entry, 259, 3, 1, 1);
        WriteEntry(output, ref entry, 262, 3, 1, 2);
        WriteEntry(output, ref entry, 277, 3, 1, 3);
        WriteEntry(output, ref entry, 284, 3, 1, 1);
        WriteEntry(output, ref entry, 322, 4, 1, 2);
        WriteEntry(output, ref entry, 323, 4, 1, 2);
        WriteEntry(output, ref entry, 324, 4, 2, offsetsOffset);
        WriteEntry(output, ref entry, 325, 4, 2, countsOffset);
        WriteUInt32(output, entry, 0);
        for (int index = 0; index < 3; index++) WriteUInt16(output, bitsOffset + index * 2, 8);
        for (int index = 0; index < 2; index++) {
            WriteUInt32(output, offsetsOffset + index * 4, pixelsOffset + index * 12);
            WriteUInt32(output, countsOffset + index * 4, 12);
        }
        WriteRgb(output, pixelsOffset, 255, 0, 0);
        WriteRgb(output, pixelsOffset + 3, 0, 255, 0);
        WriteRgb(output, pixelsOffset + 6, 255, 255, 0);
        WriteRgb(output, pixelsOffset + 9, 0, 255, 255);
        WriteRgb(output, pixelsOffset + 12, 0, 0, 255);
        WriteRgb(output, pixelsOffset + 18, 255, 255, 255);
        return output;
    }

    private static void WriteHeader(byte[] output, int ifdOffset, int entryCount) {
        output[0] = (byte)'I';
        output[1] = (byte)'I';
        WriteUInt16(output, 2, 42);
        WriteUInt32(output, 4, ifdOffset);
        WriteUInt16(output, ifdOffset, entryCount);
    }

    private static void WriteRgb(byte[] output, int offset, byte red, byte green, byte blue) {
        output[offset] = red;
        output[offset + 1] = green;
        output[offset + 2] = blue;
    }

    private static void WriteEntry(byte[] output, ref int offset, int tag, int type, int count, int value) {
        WriteUInt16(output, offset, tag);
        WriteUInt16(output, offset + 2, type);
        WriteUInt32(output, offset + 4, count);
        WriteUInt32(output, offset + 8, value);
        offset += 12;
    }

    private static void WriteUInt16(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
        output[offset + 2] = (byte)(value >> 16);
        output[offset + 3] = (byte)(value >> 24);
    }
}
