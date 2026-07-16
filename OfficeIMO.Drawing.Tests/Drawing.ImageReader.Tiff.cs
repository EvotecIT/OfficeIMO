using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OfficeImageReader_IdentifiesBigTiffMetadata(bool littleEndian) {
        byte[] data = CreateBigTiff(littleEndian, 11, 9, 300);

        bool identified = OfficeImageReader.TryIdentify(data, fileName: null, out OfficeImageInfo info);

        Assert.True(identified);
        Assert.Equal(OfficeImageFormat.Tiff, info.Format);
        Assert.Equal(11, info.Width);
        Assert.Equal(9, info.Height);
        Assert.Equal(300D, info.DpiX);
        Assert.Equal(300D, info.DpiY);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OfficeImageReader_DefaultsUnitlessBigTiffResolution(bool littleEndian) {
        byte[] data = CreateBigTiff(littleEndian, 11, 9, 300, resolutionUnit: 1);

        bool identified = OfficeImageReader.TryIdentify(data, fileName: null, out OfficeImageInfo info);

        Assert.True(identified);
        Assert.Equal(96D, info.DpiX);
        Assert.Equal(96D, info.DpiY);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OfficeImageReader_ReadsClassicTiffRationalsAsUnsignedValues(bool littleEndian) {
        byte[] data = CreateClassicTiff(
            littleEndian,
            width: 11,
            height: 9,
            xNumerator: 0x80000000u,
            xDenominator: 2,
            yNumerator: 1,
            yDenominator: 0x80000000u);

        bool identified = OfficeImageReader.TryIdentifyByContent(data, fileName: null, out OfficeImageInfo info);

        Assert.True(identified);
        Assert.Equal(1073741824D, info.DpiX);
        Assert.Equal(1D / 0x80000000u, info.DpiY);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OfficeImageReader_RejectsTruncatedClassicTiffDirectory(bool littleEndian) {
        byte[] data = CreateClassicTiff(
            littleEndian,
            width: 11,
            height: 9,
            xNumerator: 300,
            xDenominator: 1,
            yNumerator: 300,
            yDenominator: 1);
        Array.Resize(ref data, 34);

        Assert.False(OfficeImageReader.TryIdentifyByContent(data, fileName: null, out _));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OfficeImageReader_RejectsClassicTiffScalarTagsWithMultipleValues(bool littleEndian) {
        const int ifdOffset = 8;
        var data = new byte[38];
        data[0] = littleEndian ? (byte)'I' : (byte)'M';
        data[1] = data[0];
        WriteUInt16(data, 2, 42, littleEndian);
        WriteUInt32(data, 4, ifdOffset, littleEndian);
        WriteUInt16(data, ifdOffset, 2, littleEndian);
        WriteClassicLongEntry(data, 10, 256, 1, littleEndian);
        WriteUInt32(data, 14, 2, littleEndian);
        WriteClassicLongEntry(data, 22, 257, 1, littleEndian);
        WriteUInt32(data, 26, 2, littleEndian);

        Assert.False(OfficeImageReader.TryIdentifyByContent(data, fileName: null, out _));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OfficeImageReader_RejectsTruncatedBigTiffDirectory(bool littleEndian) {
        byte[] data = CreateBigTiff(littleEndian, 11, 9, 300);
        Array.Resize(ref data, 64);

        Assert.False(OfficeImageReader.TryIdentifyByContent(data, fileName: null, out _));
    }

    private static byte[] CreateClassicTiff(
        bool littleEndian,
        int width,
        int height,
        uint xNumerator,
        uint xDenominator,
        uint yNumerator,
        uint yDenominator) {
        const int ifdOffset = 8;
        const int xRationalOffset = 80;
        const int yRationalOffset = 88;
        var bytes = new byte[96];
        bytes[0] = littleEndian ? (byte)'I' : (byte)'M';
        bytes[1] = bytes[0];
        WriteUInt16(bytes, 2, 42, littleEndian);
        WriteUInt32(bytes, 4, ifdOffset, littleEndian);
        WriteUInt16(bytes, ifdOffset, 5, littleEndian);
        WriteClassicLongEntry(bytes, 10, 256, (uint)width, littleEndian);
        WriteClassicLongEntry(bytes, 22, 257, (uint)height, littleEndian);
        WriteClassicRationalEntry(bytes, 34, 282, xRationalOffset, littleEndian);
        WriteClassicRationalEntry(bytes, 46, 283, yRationalOffset, littleEndian);
        WriteClassicShortEntry(bytes, 58, 296, 2, littleEndian);
        WriteUInt32(bytes, xRationalOffset, xNumerator, littleEndian);
        WriteUInt32(bytes, xRationalOffset + 4, xDenominator, littleEndian);
        WriteUInt32(bytes, yRationalOffset, yNumerator, littleEndian);
        WriteUInt32(bytes, yRationalOffset + 4, yDenominator, littleEndian);
        return bytes;
    }

    private static void WriteClassicLongEntry(
        byte[] bytes,
        int offset,
        ushort tag,
        uint value,
        bool littleEndian) {
        WriteClassicEntryHeader(bytes, offset, tag, 4, littleEndian);
        WriteUInt32(bytes, offset + 8, value, littleEndian);
    }

    private static void WriteClassicRationalEntry(
        byte[] bytes,
        int offset,
        ushort tag,
        int valueOffset,
        bool littleEndian) {
        WriteClassicEntryHeader(bytes, offset, tag, 5, littleEndian);
        WriteUInt32(bytes, offset + 8, (uint)valueOffset, littleEndian);
    }

    private static void WriteClassicShortEntry(
        byte[] bytes,
        int offset,
        ushort tag,
        ushort value,
        bool littleEndian) {
        WriteClassicEntryHeader(bytes, offset, tag, 3, littleEndian);
        WriteUInt16(bytes, offset + 8, value, littleEndian);
    }

    private static void WriteClassicEntryHeader(
        byte[] bytes,
        int offset,
        ushort tag,
        ushort type,
        bool littleEndian) {
        WriteUInt16(bytes, offset, tag, littleEndian);
        WriteUInt16(bytes, offset + 2, type, littleEndian);
        WriteUInt32(bytes, offset + 4, 1, littleEndian);
    }

    private static byte[] CreateBigTiff(
        bool littleEndian,
        int width,
        int height,
        uint dpi,
        ushort resolutionUnit = 2) {
        var bytes = new byte[132];
        bytes[0] = littleEndian ? (byte)'I' : (byte)'M';
        bytes[1] = bytes[0];
        WriteUInt16(bytes, 2, 43, littleEndian);
        WriteUInt16(bytes, 4, 8, littleEndian);
        WriteUInt64(bytes, 8, 16, littleEndian);
        WriteUInt64(bytes, 16, 5, littleEndian);
        WriteLongEntry(bytes, 24, 256, (uint)width, littleEndian);
        WriteLongEntry(bytes, 44, 257, (uint)height, littleEndian);
        WriteRationalEntry(bytes, 64, 282, dpi, 1, littleEndian);
        WriteRationalEntry(bytes, 84, 283, dpi, 1, littleEndian);
        WriteShortEntry(bytes, 104, 296, resolutionUnit, littleEndian);
        return bytes;
    }

    private static void WriteLongEntry(byte[] bytes, int offset, ushort tag, uint value, bool littleEndian) {
        WriteEntryHeader(bytes, offset, tag, 4, littleEndian);
        WriteUInt32(bytes, offset + 12, value, littleEndian);
    }

    private static void WriteRationalEntry(
        byte[] bytes,
        int offset,
        ushort tag,
        uint numerator,
        uint denominator,
        bool littleEndian) {
        WriteEntryHeader(bytes, offset, tag, 5, littleEndian);
        WriteUInt32(bytes, offset + 12, numerator, littleEndian);
        WriteUInt32(bytes, offset + 16, denominator, littleEndian);
    }

    private static void WriteShortEntry(byte[] bytes, int offset, ushort tag, ushort value, bool littleEndian) {
        WriteEntryHeader(bytes, offset, tag, 3, littleEndian);
        WriteUInt16(bytes, offset + 12, value, littleEndian);
    }

    private static void WriteEntryHeader(byte[] bytes, int offset, ushort tag, ushort type, bool littleEndian) {
        WriteUInt16(bytes, offset, tag, littleEndian);
        WriteUInt16(bytes, offset + 2, type, littleEndian);
        WriteUInt64(bytes, offset + 4, 1, littleEndian);
    }

    private static void WriteUInt16(byte[] bytes, int offset, ushort value, bool littleEndian) {
        int first = littleEndian ? 0 : 8;
        int second = littleEndian ? 8 : 0;
        bytes[offset] = (byte)(value >> first);
        bytes[offset + 1] = (byte)(value >> second);
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value, bool littleEndian) {
        for (int index = 0; index < 4; index++) {
            int shift = (littleEndian ? index : 3 - index) * 8;
            bytes[offset + index] = (byte)(value >> shift);
        }
    }

    private static void WriteUInt64(byte[] bytes, int offset, ulong value, bool littleEndian) {
        for (int index = 0; index < 8; index++) {
            int shift = (littleEndian ? index : 7 - index) * 8;
            bytes[offset + index] = (byte)(value >> shift);
        }
    }
}
