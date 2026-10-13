using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class ImageMetadataTypedArrayTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExifNumericArraysRetainTheirTypesAndValuesThroughPublicProfileEdits(bool littleEndian) {
        byte[] empty = littleEndian
            ? [73, 73, 42, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0]
            : [77, 77, 0, 42, 0, 0, 0, 8, 0, 0, 0, 0, 0, 0];
        var metadata = OfficeImageMetadata.ParseExifProfile(empty);
        (OfficeExifDataType Type, Array Values)[] fields = [
            (OfficeExifDataType.SignedByte, new sbyte[] { -2, 3 }),
            (OfficeExifDataType.Short, new ushort[] { 2, 65535 }),
            (OfficeExifDataType.SignedShort, new short[] { -32768, 3 }),
            (OfficeExifDataType.Long, new uint[] { 2, uint.MaxValue }),
            (OfficeExifDataType.SignedLong, new int[] { int.MinValue, 3 }),
            (OfficeExifDataType.Rational, new OfficeRational[] { new(2, 3), new(5, 7) }),
            (OfficeExifDataType.SignedRational, new OfficeSignedRational[] { new(-2, 3), new(5, 7) }),
            (OfficeExifDataType.Float, new float[] { -1.25f, 2.5f }),
            (OfficeExifDataType.Double, new double[] { -1.25, 2.5 })
        ];
        for (int i = 0; i < fields.Length; i++)
            metadata.SetExifValue(new OfficeExifTag((ushort)(65000 + i), fields[i].Type, OfficeExifDirectory.Image), fields[i].Values);

        var reopened = OfficeImageMetadata.ParseExifProfile(metadata.EncodeExifProfile()!);

        for (int i = 0; i < fields.Length; i++) {
            object value = reopened.GetExifValue(new OfficeExifTag((ushort)(65000 + i), fields[i].Type, OfficeExifDirectory.Image))!.Value;
            Assert.Equal(fields[i].Values.GetType(), value.GetType());
            Assert.Equal(fields[i].Values.Cast<object>(), ((Array)value).Cast<object>());
        }
    }
}
