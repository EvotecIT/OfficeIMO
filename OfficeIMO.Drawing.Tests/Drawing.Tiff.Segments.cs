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

    [Fact]
    public void TiffInventoryAcceptsAnIgnorableIfdTypedSubIfdPointer() {
        byte[] tiff = CreateTiledRgbTiff(includeSubIfd: true);

        Assert.True(OfficeRasterContainerInspector.TryInspect(
            tiff, out OfficeRasterContainerInfo? container));
        Assert.Single(container!.Frames);
        Assert.True(OfficeTiffCodec.TryDecode(tiff, out OfficeRasterImage? image));
        Assert.Equal((3, 2), (image!.Width, image.Height));
    }

    [Fact]
    public void TiffTileDecodeRejectsAnAggregateWorkingSetAboveTheManagedLimit() {
        const int oneHundredTwentyEightMiB = 128 * 1024 * 1024;

        Assert.False(OfficeTiffCodec.IsTiffDecodeWorkingSetWithinLimit(
            encodedLength: oneHundredTwentyEightMiB,
            sourceLength: oneHundredTwentyEightMiB,
            scratchLength: oneHundredTwentyEightMiB,
            finalRgbaLength: oneHundredTwentyEightMiB,
            segmentMetadataBytes: 8L,
            retainPixels: true,
            compression: (int)OfficeTiffCompression.Deflate,
            maximumCompressedSegmentLength: oneHundredTwentyEightMiB,
            maximumDecodedSegmentLength: oneHundredTwentyEightMiB));
        Assert.True(OfficeTiffCodec.IsTiffDecodeWorkingSetWithinLimit(
            encodedLength: 1024,
            sourceLength: 24,
            scratchLength: 12,
            finalRgbaLength: 32,
            segmentMetadataBytes: 16L,
            retainPixels: true,
            compression: (int)OfficeTiffCompression.Deflate,
            maximumCompressedSegmentLength: 128,
            maximumDecodedSegmentLength: 12));
    }

    [Fact]
    public void TiffSelectedDecodeRejectsAliasedStripWorkBeforeDecoding() {
        byte[] tiff = CreateAliasedPackBitsStripTiff();

        Assert.False(OfficeTiffCodec.TryDecode(tiff, out _));
    }

    [Fact]
    public void TiffSelectedDecodeRejectsAggregatePaddedTileWorkBeforeAllocatingScratch() {
        const int tileDecodedBytes = 44_739_242 * 3;
        Assert.True(tileDecodedBytes <= OfficeRasterGuards.MaximumDecodedBytes);
        Assert.True(2L * (tileDecodedBytes + 3L) > OfficeRasterGuards.MaximumDecodedBytes);

        byte[] tiff = CreateOversizedPaddedTileWorkTiff();

        Assert.False(OfficeTiffCodec.TryDecode(tiff, out _));
    }

    [Fact]
    public void TiffMultiPageEncodeRejectsTheCombinedSourceStripAndOutputPeak() {
        const int oneHundredTwentyFourMiB = 124 * 1024 * 1024;
        const int oneHundredTwentyFiveMiB = 125 * 1024 * 1024;

        Assert.False(OfficeTiffCodec.IsMultiPageTiffWorkingSetWithinLimit(
            sourceBytes: oneHundredTwentyFourMiB,
            retainedStripBytes: oneHundredTwentyFiveMiB,
            outputBytes: oneHundredTwentyFiveMiB));
        Assert.True(OfficeTiffCodec.IsMultiPageTiffWorkingSetWithinLimit(
            sourceBytes: 1024,
            retainedStripBytes: 512,
            outputBytes: 768));
        Assert.False(OfficeTiffCodec.CanBeginMultiPageStripEncoding(
            sourceBytes: 200L * 1024L * 1024L,
            retainedStripBytes: 0L,
            pendingAllocationBytes: 200L * 1024L * 1024L));
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

    private static byte[] CreateTiledRgbTiff(bool includeSubIfd = false) {
        int entryCount = includeSubIfd ? 12 : 11;
        const int ifdOffset = 8;
        int dataOffset = ifdOffset + 2 + entryCount * 12 + 4;
        int bitsOffset = dataOffset;
        int offsetsOffset = bitsOffset + 6;
        int countsOffset = offsetsOffset + 8;
        int pixelsOffset = countsOffset + 8;
        int subIfdOffset = pixelsOffset + 24;
        var output = new byte[subIfdOffset + (includeSubIfd ? 6 : 0)];
        WriteHeader(output, ifdOffset, entryCount);
        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, 3);
        WriteEntry(output, ref entry, 257, 4, 1, 2);
        WriteEntry(output, ref entry, 258, 3, 3, bitsOffset);
        WriteEntry(output, ref entry, 259, 3, 1, 1);
        WriteEntry(output, ref entry, 262, 3, 1, 2);
        WriteEntry(output, ref entry, 277, 3, 1, 3);
        WriteEntry(output, ref entry, 284, 3, 1, 1);
        if (includeSubIfd) WriteEntry(output, ref entry, 330, 13, 1, subIfdOffset);
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
        if (includeSubIfd) {
            WriteUInt16(output, subIfdOffset, 0);
            WriteUInt32(output, subIfdOffset + 2, 0);
        }
        return output;
    }

    private static byte[] CreateAliasedPackBitsStripTiff() {
        const int stripCount = 256;
        const int payloadLength = 1024 * 1024;
        const int entryCount = 8;
        const int ifdOffset = 8;
        int dataOffset = ifdOffset + 2 + entryCount * 12 + 4;
        int offsetsOffset = dataOffset;
        int countsOffset = offsetsOffset + stripCount * 4;
        int payloadOffset = countsOffset + stripCount * 4;
        var output = new byte[payloadOffset + payloadLength];
        WriteHeader(output, ifdOffset, entryCount);
        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, 1);
        WriteEntry(output, ref entry, 257, 4, 1, stripCount);
        WriteEntry(output, ref entry, 258, 3, 1, 8);
        WriteEntry(output, ref entry, 259, 3, 1, (int)OfficeTiffCompression.PackBits);
        WriteEntry(output, ref entry, 262, 3, 1, 1);
        WriteEntry(output, ref entry, 273, 4, stripCount, offsetsOffset);
        WriteEntry(output, ref entry, 278, 4, 1, 1);
        WriteEntry(output, ref entry, 279, 4, stripCount, countsOffset);
        WriteUInt32(output, entry, 0);
        for (int index = 0; index < stripCount; index++) {
            WriteUInt32(output, offsetsOffset + index * 4, payloadOffset);
            WriteUInt32(output, countsOffset + index * 4, payloadLength);
        }
        output[payloadOffset] = 0;
        output[payloadOffset + 1] = 127;
        for (int index = payloadOffset + 2; index < output.Length; index++) output[index] = 0x80;
        return output;
    }

    private static byte[] CreateOversizedPaddedTileWorkTiff() {
        const int entryCount = 11;
        const int ifdOffset = 8;
        const int tileWidth = 44_739_242;
        int dataOffset = ifdOffset + 2 + entryCount * 12 + 4;
        int bitsOffset = dataOffset;
        int offsetsOffset = bitsOffset + 6;
        int countsOffset = offsetsOffset + 8;
        int payloadOffset = countsOffset + 8;
        var output = new byte[payloadOffset + 3];
        WriteHeader(output, ifdOffset, entryCount);
        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, 1);
        WriteEntry(output, ref entry, 257, 4, 1, 2);
        WriteEntry(output, ref entry, 258, 3, 3, bitsOffset);
        WriteEntry(output, ref entry, 259, 3, 1, (int)OfficeTiffCompression.PackBits);
        WriteEntry(output, ref entry, 262, 3, 1, 2);
        WriteEntry(output, ref entry, 277, 3, 1, 3);
        WriteEntry(output, ref entry, 284, 3, 1, 1);
        WriteEntry(output, ref entry, 322, 4, 1, tileWidth);
        WriteEntry(output, ref entry, 323, 4, 1, 1);
        WriteEntry(output, ref entry, 324, 4, 2, offsetsOffset);
        WriteEntry(output, ref entry, 325, 4, 2, countsOffset);
        WriteUInt32(output, entry, 0);
        for (int index = 0; index < 3; index++) WriteUInt16(output, bitsOffset + index * 2, 8);
        for (int index = 0; index < 2; index++) {
            WriteUInt32(output, offsetsOffset + index * 4, payloadOffset);
            WriteUInt32(output, countsOffset + index * 4, 3);
        }
        output[payloadOffset] = 0;
        output[payloadOffset + 1] = 127;
        output[payloadOffset + 2] = 0x80;
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
