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

    private static byte[] CreateBigTiff(bool littleEndian, int width, int height, uint dpi) {
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
        WriteShortEntry(bytes, 104, 296, 2, littleEndian);
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
