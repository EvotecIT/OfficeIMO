using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class ExifTypedDecodingTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NumericFieldsKeepTheirPublicScalarAndArrayTypes(bool littleEndian) {
        Verify(OfficeExifDataType.SignedByte, littleEndian, 1,
            new byte[] { 254, 127 }, new sbyte[] { -2, 127 });
        Verify(OfficeExifDataType.Short, littleEndian, 2,
            new byte[] { 0x34, 0x12, 0xdc, 0xfe }, new ushort[] { 0x1234, 0xfedc });
        Verify(OfficeExifDataType.SignedShort, littleEndian, 2,
            new byte[] { 254, 255, 0x34, 0x12 }, new short[] { -2, 0x1234 });
        Verify(OfficeExifDataType.Long, littleEndian, 4,
            new byte[] { 0x78, 0x56, 0x34, 0x12, 0x98, 0xba, 0xdc, 0xfe }, new uint[] { 0x12345678, 0xfedcba98 });
        Verify(OfficeExifDataType.SignedLong, littleEndian, 4,
            new byte[] { 254, 255, 255, 255, 0x78, 0x56, 0x34, 0x12 }, new int[] { -2, 0x12345678 });
        Verify(OfficeExifDataType.Rational, littleEndian, 4,
            new byte[] { 1, 0, 0, 0, 3, 0, 0, 0, 7, 0, 0, 0, 2, 0, 0, 0 },
            new[] { new OfficeRational(1, 3), new OfficeRational(7, 2) });
        Verify(OfficeExifDataType.SignedRational, littleEndian, 4,
            new byte[] { 255, 255, 255, 255, 3, 0, 0, 0, 7, 0, 0, 0, 254, 255, 255, 255 },
            new[] { new OfficeSignedRational(-1, 3), new OfficeSignedRational(7, -2) });
        Verify(OfficeExifDataType.Float, littleEndian, 4,
            new byte[] { 0, 0, 0xc0, 0xbf, 0, 0, 0x10, 0x40 }, new float[] { -1.5F, 2.25F });
        Verify(OfficeExifDataType.Double, littleEndian, 8,
            new byte[] { 0, 0, 0, 0, 0, 0, 0xf8, 0xbf, 0, 0, 0, 0, 0, 0, 2, 0x40 }, new double[] { -1.5D, 2.25D });
    }

    private static void Verify<T>(OfficeExifDataType type, bool littleEndian, int byteOrderUnit,
        byte[] littleEndianPayload, T[] expected) {
        // Hand-authored TIFF fields avoid using the encoder as the decoding oracle.
        byte[] payload = (byte[])littleEndianPayload.Clone();
        if (!littleEndian) {
            for (int offset = 0; offset < payload.Length; offset += byteOrderUnit) {
                Array.Reverse(payload, offset, byteOrderUnit);
            }
        }
        var tag = new OfficeExifTag(65000, type, OfficeExifDirectory.Image);
        byte[] scalarProfile = CreateProfile(type, 1, payload, payload.Length / expected.Length, littleEndian);
        object scalar = OfficeImageMetadata.ParseExifProfile(scalarProfile).GetExifValue(tag)!.Value;
        Assert.Equal(expected[0], Assert.IsType<T>(scalar));

        byte[] arrayProfile = CreateProfile(type, expected.Length, payload, payload.Length, littleEndian);
        object array = OfficeImageMetadata.ParseExifProfile(arrayProfile).GetExifValue(tag)!.Value;
        Assert.Equal(expected, Assert.IsType<T[]>(array));
    }

    private static byte[] CreateProfile(OfficeExifDataType type, int count, byte[] payload, int length, bool little) {
        var profile = new byte[26 + (length > 4 ? length : 0)];
        profile[0] = profile[1] = little ? (byte)73 : (byte)77;
        Put(profile, 2, 42, 2, little);
        Put(profile, 4, 8, 4, little);
        Put(profile, 8, 1, 2, little);
        Put(profile, 10, 65000, 2, little);
        Put(profile, 12, (uint)type, 2, little);
        Put(profile, 14, (uint)count, 4, little);
        if (length > 4) Put(profile, 18, 26, 4, little);
        Buffer.BlockCopy(payload, 0, profile, length > 4 ? 26 : 18, length);
        return profile;
    }

    private static void Put(byte[] bytes, int offset, uint value, int length, bool little) {
        for (int i = 0; i < length; i++) {
            bytes[offset + (little ? i : length - 1 - i)] = (byte)(value >> (i * 8));
        }
    }
}
