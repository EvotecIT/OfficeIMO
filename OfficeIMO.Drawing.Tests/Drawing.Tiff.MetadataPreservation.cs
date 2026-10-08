using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class TiffMetadataPreservationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ParsedAsciiListsAndZeroCountArraysSurviveIndependentSnapshotsAndUnrelatedEdits(bool bigEndian, bool embeddedNul) {
        byte[] encoded = MakeMetadataFixture(bigEndian, embeddedNul);
        string expectedText = embeddedNul ? "First\0Second" : "First";
        Assert.True(OfficeTiffStructureValidator.TryValidate(encoded, 0, encoded.Length));
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? before));
        OfficeImageMetadata snapshot = OfficeImageMetadata.Read(encoded);
        Assert.Equal(expectedText, snapshot.GetExifValue(OfficeExifTag.Copyright)!.Value);
        for (int type = 1; type <= 12; type++) {
            object value = snapshot.GetExifValue(new OfficeExifTag((ushort)(65000 + type), (OfficeExifDataType)type, OfficeExifDirectory.Exif))!.Value;
            if (type == 2) Assert.Equal("", value);
            else Assert.Empty(Assert.IsAssignableFrom<Array>(value));
        }
        Assert.Empty(Assert.IsType<OfficeRational[]>(snapshot.GetExifValue(OfficeExifTag.GPSLatitude)!.Value));
        snapshot.SetExifValue(OfficeExifTag.Artist, "Edited creator");
        byte[] applied = OfficeImageMetadata.Apply(encoded, snapshot);
        AssertOriginalFields(encoded, applied);
        Assert.True(OfficeRasterImageDecoder.TryDecode(applied, out OfficeRasterImage? after));
        Assert.Equal(before!.GetPixels(), after!.GetPixels());
        OfficeImageMetadata parsedProfile = OfficeImageMetadata.ParseExifProfile(encoded);
        byte[] fromParsedProfile = OfficeImageMetadata.Apply(OfficeTiffCodec.Encode(new OfficeRasterImage(2, 1, OfficeColor.Blue)), parsedProfile);
        OfficeImageMetadata relocatedProfile = OfficeImageMetadata.Read(fromParsedProfile);
        Assert.Equal(new ushort[] { 0x1234, 0xABCD }, Assert.IsType<ushort[]>(relocatedProfile.GetExifValue(new OfficeExifTag(65020, OfficeExifDataType.Short, OfficeExifDirectory.Exif))!.Value));
        OfficeRational relocatedRational = Assert.IsType<OfficeRational>(relocatedProfile.GetExifValue(new OfficeExifTag(65021, OfficeExifDataType.Rational, OfficeExifDirectory.Exif))!.Value);
        Assert.Equal(0x11223344U, relocatedRational.Numerator); Assert.Equal(0x55667788U, relocatedRational.Denominator);
        OfficeImageMetadata clone = snapshot.Clone();
        Array.Clear(encoded, 0, encoded.Length);
        byte[] exported = clone.EncodeExifProfile()!;
        OfficeImageMetadata exportedMetadata = OfficeImageMetadata.ParseExifProfile(exported);
        Assert.Equal(expectedText, exportedMetadata.GetExifValue(OfficeExifTag.Copyright)!.Value);
        foreach (int type in new[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 }) Assert.Equal(0U, ReadField(exported, OfficeExifDirectory.Exif, (ushort)(65000 + type)).Count);
        Assert.Equal(0U, ReadField(exported, OfficeExifDirectory.Gps, 2).Count);
        Assert.Equal(Encoding.ASCII.GetBytes(expectedText + "\0\0"), ReadField(exported, OfficeExifDirectory.Image, 33432).Bytes);
        Assert.Equal(new ushort[] { 0x1234, 0xABCD }, Assert.IsType<ushort[]>(exportedMetadata.GetExifValue(new OfficeExifTag(65020, OfficeExifDataType.Short, OfficeExifDirectory.Exif))!.Value));
        OfficeRational rational = Assert.IsType<OfficeRational>(exportedMetadata.GetExifValue(new OfficeExifTag(65021, OfficeExifDataType.Rational, OfficeExifDirectory.Exif))!.Value);
        Assert.Equal(0x11223344U, rational.Numerator); Assert.Equal(0x55667788U, rational.Denominator);
        Assert.Equal("Edited creator", exportedMetadata.GetExifValue(OfficeExifTag.Artist)!.Value);
        // The same parsed fields must also survive moving from a classic Exif profile
        // into a freshly encoded image whose byte order differs from the original.
        byte[] transferred = OfficeImageMetadata.Apply(OfficeTiffCodec.Encode(new OfficeRasterImage(2, 1, OfficeColor.Blue)), exportedMetadata);
        Assert.Equal(expectedText, OfficeImageMetadata.Read(transferred).GetExifValue(OfficeExifTag.Copyright)!.Value);
        Assert.Equal(0U, ReadField(transferred, OfficeExifDirectory.Gps, 2).Count);
        byte[] stripped = OfficeImageMetadata.Remove(applied, OfficeImageMetadataProfileKinds.Exif).EncodedBytes;
        OfficeImageMetadata remaining = OfficeImageMetadata.Read(stripped);
        Assert.Null(remaining.GetExifValue(OfficeExifTag.Copyright));
        Assert.Null(remaining.GetExifValue(OfficeExifTag.GPSLatitude));
        Assert.DoesNotContain("First", Encoding.ASCII.GetString(stripped));
        Assert.True(OfficeRasterImageDecoder.TryDecode(stripped, out OfficeRasterImage? removed));
        Assert.Equal(before.GetPixels(), removed!.GetPixels());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(6)]
    [InlineData(8)]
    public void ImplicitExifRemovalPreservesRenderingOrientationOnEveryTiffPage(int orientation) {
        var first = new OfficeRasterImage(3, 2, OfficeColor.Red); first.SetPixel(1, 0, OfficeColor.Blue);
        var second = new OfficeRasterImage(2, 3, OfficeColor.Green); second.SetPixel(0, 1, OfficeColor.Red);
        byte[] encoded = OfficeTiffCodec.EncodePages(new[] { first, second });
        foreach (int page in Pages(encoded)) {
            int entry = FindEntry(encoded, page, 274); Put(encoded, entry + 8, (uint)orientation, 2, true);
        }
        var metadata = OfficeImageMetadata.Read(encoded); metadata.SetExifValue(OfficeExifTag.Artist, "Private creator");
        encoded = OfficeImageMetadata.Apply(encoded, metadata);
        metadata = OfficeImageMetadata.Read(encoded); metadata.ClearExif();
        byte[] appliedClear = OfficeImageMetadata.Apply(encoded, metadata);
        byte[] removed = OfficeImageMetadata.Remove(encoded, OfficeImageMetadataProfileKinds.Exif).EncodedBytes;
        byte[] removedAll = OfficeImageMetadata.Remove(encoded, OfficeImageMetadataProfileKinds.All).EncodedBytes;
        foreach (byte[] result in new[] { appliedClear, removed, removedAll }) {
            Assert.DoesNotContain("Private creator", Encoding.ASCII.GetString(result));
            List<int> pages = Pages(result); Assert.Equal(2, pages.Count);
            for (int index = 0; index < pages.Count; index++) {
                Assert.Equal((uint)orientation, Read(result, FindEntry(result, pages[index], 274) + 8, 2, true));
                var options = new OfficeRasterDecodeOptions { FrameIndex = index };
                Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, options, out OfficeRasterImage? before, out _));
                Assert.True(OfficeRasterImageDecoder.TryDecode(result, options, out OfficeRasterImage? after, out _));
                Assert.Equal(before!.Width, after!.Width); Assert.Equal(before.Height, after.Height); Assert.Equal(before.GetPixels(), after.GetPixels());
                Assert.Equal(PixelPayload(encoded, Pages(encoded)[index]), PixelPayload(result, pages[index]));
            }
        }
    }

    [Fact]
    public void ExplicitOrientationEditsRemainAvailableAndStrictUserValidationIsUnchanged() {
        byte[] encoded = OfficeTiffCodec.Encode(new OfficeRasterImage(3, 2, OfficeColor.Red));
        var metadata = OfficeImageMetadata.Read(encoded);
        Assert.Throws<ArgumentException>(() => metadata.SetExifValue(OfficeExifTag.Copyright, "First\0Second"));
        Assert.Throws<ArgumentException>(() => metadata.SetExifValue(OfficeExifTag.UserComment, Array.Empty<byte>()));
        Assert.Throws<ArgumentException>(() => metadata.SetExifValue(OfficeExifTag.GPSLatitude, Array.Empty<OfficeRational>()));
        metadata.SetExifValue(OfficeExifTag.Orientation, (ushort)6);
        encoded = OfficeImageMetadata.Apply(encoded, metadata);
        Assert.Equal((ushort)6, OfficeImageMetadata.Read(encoded).GetExifValue(OfficeExifTag.Orientation)!.Value);
        metadata = OfficeImageMetadata.Read(encoded); Assert.True(metadata.RemoveExifValue(OfficeExifTag.Orientation));
        encoded = OfficeImageMetadata.Apply(encoded, metadata);
        Assert.Null(OfficeImageMetadata.Read(encoded).GetExifValue(OfficeExifTag.Orientation));
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? image)); Assert.Equal(3, image!.Width); Assert.Equal(2, image.Height);
    }

    private static byte[] MakeMetadataFixture(bool bigEndian, bool embeddedNul) {
        // A tiny uncompressed grayscale TIFF assembled independently of the metadata writer.
        bool little = !bigEndian;
        const int root = 8, rootCount = 12;
        const int exif = root + 6 + rootCount * 12, exifCount = 14;
        const int gps = exif + 6 + exifCount * 12;
        const int text = gps + 18;
        byte[] ascii = Encoding.ASCII.GetBytes((embeddedNul ? "First\0Second" : "First") + "\0\0");
        int rational = text + ascii.Length;
        int pixels = rational + 8;
        var bytes = new byte[pixels + 6]; bytes[0] = bytes[1] = (byte)(little ? 73 : 77);
        Put(bytes, 2, 42, 2, little); Put(bytes, 4, root, 4, little); Put(bytes, root, rootCount, 2, little);
        int at = root + 2;
        Entry(256, 4, 1, 3); Entry(257, 4, 1, 2); Entry(258, 3, 1, 8); Entry(259, 3, 1, 1); Entry(262, 3, 1, 1);
        Entry(273, 4, 1, (uint)pixels); Entry(274, 3, 1, 1); Entry(277, 3, 1, 1); Entry(278, 4, 1, 2); Entry(279, 4, 1, 6);
        Entry(33432, 2, (uint)ascii.Length, (uint)text); Entry(34665, 4, 1, exif);
        // GPS pointer is added as a thirteenth root entry in a replacement table after the payload.
        Put(bytes, exif, exifCount, 2, little); at = exif + 2;
        for (int type = 1; type <= 12; type++) Entry((uint)(65000 + type), (uint)type, 0, 0);
        int shorts = at; Entry(65020, 3, 2, 0); Put(bytes, shorts + 8, 0x1234, 2, little); Put(bytes, shorts + 10, 0xABCD, 2, little);
        Entry(65021, 5, 1, (uint)rational); Put(bytes, rational, 0x11223344, 4, little); Put(bytes, rational + 4, 0x55667788, 4, little);
        Put(bytes, gps, 1, 2, little); at = gps + 2; Entry(2, 5, 0, 0);
        Buffer.BlockCopy(ascii, 0, bytes, text, ascii.Length); for (int i = 0; i < 6; i++) bytes[pixels + i] = (byte)(20 + 30 * i);
        int newRoot = bytes.Length; Array.Resize(ref bytes, newRoot + 6 + (rootCount + 1) * 12);
        Put(bytes, 4, (uint)newRoot, 4, little); Put(bytes, newRoot, rootCount + 1, 2, little); Buffer.BlockCopy(bytes, root + 2, bytes, newRoot + 2, rootCount * 12);
        at = newRoot + 2 + rootCount * 12; Entry(34853, 4, 1, gps);
        return bytes;
        void Entry(uint tag, uint type, uint count, uint value) { Put(bytes, at, tag, 2, little); Put(bytes, at + 2, type, 2, little); Put(bytes, at + 4, count, 4, little); Put(bytes, at + 8, value, type == 3 && count == 1 ? 2 : 4, little); at += 12; }
    }

    private static void AssertOriginalFields(byte[] original, byte[] rewritten) {
        Assert.Equal(ReadField(original, OfficeExifDirectory.Image, 33432).Bytes, ReadField(rewritten, OfficeExifDirectory.Image, 33432).Bytes);
        for (int type = 1; type <= 12; type++) Assert.Equal(0U, ReadField(rewritten, OfficeExifDirectory.Exif, (ushort)(65000 + type)).Count);
        Assert.Equal(0U, ReadField(rewritten, OfficeExifDirectory.Gps, 2).Count);
    }
    private static (uint Count, byte[] Bytes) ReadField(byte[] bytes, OfficeExifDirectory directory, ushort tag) {
        bool little = bytes[0] == 73; int root = (int)Read(bytes, 4, 4, little);
        if (directory != OfficeExifDirectory.Image) root = (int)Read(bytes, FindEntry(bytes, root, directory == OfficeExifDirectory.Exif ? 34665 : 34853) + 8, 4, little);
        int entry = FindEntry(bytes, root, tag); uint count = Read(bytes, entry + 4, 4, little); uint type = Read(bytes, entry + 2, 2, little);
        int length = checked((int)count * (type == 3 || type == 8 ? 2 : type == 4 || type == 9 || type == 11 ? 4 : type == 5 || type == 10 || type == 12 ? 8 : 1));
        int value = length <= 4 ? entry + 8 : (int)Read(bytes, entry + 8, 4, little);
        var payload = new byte[length]; Buffer.BlockCopy(bytes, value, payload, 0, length); return (count, payload);
    }
    private static List<int> Pages(byte[] bytes) {
        bool little = bytes[0] == 73; var result = new List<int>(); int at = (int)Read(bytes, 4, 4, little);
        while (at != 0) { result.Add(at); int count = (int)Read(bytes, at, 2, little); at = (int)Read(bytes, at + 2 + count * 12, 4, little); } return result;
    }
    private static byte[] PixelPayload(byte[] bytes, int page) {
        bool little = bytes[0] == 73; int at = (int)Read(bytes, FindEntry(bytes, page, 273) + 8, 4, little); int length = (int)Read(bytes, FindEntry(bytes, page, 279) + 8, 4, little);
        var payload = new byte[length]; Buffer.BlockCopy(bytes, at, payload, 0, length); return payload;
    }
    private static int FindEntry(byte[] bytes, int table, int tag) {
        bool little = bytes[0] == 73; int count = (int)Read(bytes, table, 2, little);
        for (int i = 0; i < count; i++) { int entry = table + 2 + i * 12; if (Read(bytes, entry, 2, little) == tag) return entry; } throw new InvalidDataException("The fixture field is absent.");
    }
    private static uint Read(byte[] bytes, int at, int size, bool little) { uint value = 0; for (int i = 0; i < size; i++) value |= (uint)bytes[at + i] << (8 * (little ? i : size - i - 1)); return value; }
    private static void Put(byte[] bytes, int at, uint value, int size, bool little) { for (int i = 0; i < size; i++) bytes[at + i] = (byte)(value >> (8 * (little ? i : size - i - 1))); }
}
