using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingHeifMetadataTests {
    [Theory]
    [InlineData("gzip", 0)]
    [InlineData("utf-8", 0)]
    [InlineData("", 1)]
    public void UnsupportedXmpDeclarationsRemainOpaqueToRequestedReadsAndWrites(string encoding, int protection) {
        byte[] packet;
        using (var output = new MemoryStream()) {
            using (var gzip = new GZipStream(output, CompressionMode.Compress, leaveOpen: true)) {
                byte[] text = Encoding.UTF8.GetBytes("<xmp>Original</xmp>");
                gzip.Write(text, 0, text.Length);
            }
            packet = output.ToArray();
        }
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "unused", 0, 0,
            encoding, xmpProtection: (ushort)protection, xmpBytes: packet);
        Assert.True(OfficeHeifMetadataReader.HasXmpItem(bytes));
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Equal(encoding, info!.XmpItem!.ContentEncoding);
        Assert.Equal((ushort)protection, info.XmpItem.ItemProtectionIndex);
        Assert.False(OfficeHeifMetadataReader.TryReadXmp(bytes, out string? read));
        Assert.Null(read);
        byte[] original = (byte[])bytes.Clone();
        foreach (string? replacement in new string?[] { "changed", null }) {
            Assert.False(OfficeHeifMetadataReader.TryWriteXmp(bytes, replacement, out byte[]? output));
            Assert.Null(output);
            WithFiles((source, target) => {
                File.WriteAllBytes(source, bytes);
                byte[] sentinel = { 123, 45, 67 };
                File.WriteAllBytes(target, sentinel);
                Assert.False(OfficeHeifMetadataReader.TryWriteXmp(source, target, replacement));
                Assert.Equal(sentinel, File.ReadAllBytes(target));
            });
        }
        Assert.Equal(original, bytes);
        var metadata = new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Software, "Changed");
        Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, metadata, out byte[]? edited));
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(edited!, out OfficeHeifImageInfo? editedInfo));
        Assert.Equal(encoding, editedInfo!.XmpItem!.ContentEncoding);
        Assert.Equal((ushort)protection, editedInfo.XmpItem.ItemProtectionIndex);
        OfficeHeifItemExtentInfo extent = editedInfo.XmpItem.Location!.Extents[0];
        Assert.Equal(packet, edited!.Skip(extent.Offset).Take(extent.Length).ToArray());
    }

    [Fact]
    public void ProtectedExifRemainsDeclaredAndUntouchedByIndependentXmpEdits() {
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0, exifProtection: 1);
        Assert.True(OfficeHeifMetadataReader.HasExifItem(bytes));
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Equal((ushort)1, info!.ExifItem!.ItemProtectionIndex);
        Assert.False(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out OfficeImageMetadata? read));
        Assert.Null(read);
        byte[] original = (byte[])bytes.Clone();
        var metadata = new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Software, "Changed");
        foreach (OfficeImageMetadata? replacement in new OfficeImageMetadata?[] { metadata, null }) {
            Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, replacement, out byte[]? output));
            Assert.Null(output);
            WithFiles((source, target) => {
                File.WriteAllBytes(source, bytes);
                byte[] sentinel = { 123, 45, 67 };
                File.WriteAllBytes(target, sentinel);
                Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(source, target, replacement));
                Assert.Equal(sentinel, File.ReadAllBytes(target));
            });
        }
        Assert.Equal(original, bytes);
        Assert.True(OfficeHeifMetadataReader.TryWriteXmp(bytes, "changed-XMP", out byte[]? edited));
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(edited!, out OfficeHeifImageInfo? editedInfo));
        Assert.Equal((ushort)1, editedInfo!.ExifItem!.ItemProtectionIndex);
        OfficeHeifItemExtentInfo before = info.ExifItem.Location!.Extents[0];
        OfficeHeifItemExtentInfo after = editedInfo.ExifItem.Location!.Extents[0];
        Assert.Equal(original.Skip(before.Offset).Take(before.Length).ToArray(),
            edited!.Skip(after.Offset).Take(after.Length).ToArray());
    }
}
