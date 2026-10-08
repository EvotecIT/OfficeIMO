using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingHeifMetadataTests {
    [Theory]
    [InlineData(new byte[] { 0x80 })]
    [InlineData(new byte[] { 0xC2 })]
    [InlineData(new byte[] { 0xC0, 0xAF })]
    [InlineData(new byte[] { 0xED, 0xA0, 0x80 })]
    [InlineData(new byte[] { 0xF4, 0x90, 0x80, 0x80 })]
    public void MalformedUtf8XmpRejectsByteFileAndStreamReads(byte[] packet) {
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "unused", 0, 0, xmpBytes: packet);
        byte[] original = (byte[])bytes.Clone();
        Assert.True(OfficeHeifMetadataReader.HasXmpItem(bytes));
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Equal(string.Empty, info!.XmpItem!.ContentEncoding);
        Assert.False(OfficeHeifMetadataReader.TryReadXmp(bytes, out string? text));
        Assert.Null(text);
        using var stream = new MemoryStream();
        stream.WriteByte(42);
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 1;
        Assert.False(OfficeHeifMetadataReader.TryReadXmp(stream, out text));
        Assert.Null(text);
        Assert.Equal(1L, stream.Position);
        Assert.True(stream.CanRead);
        WithFiles((source, _) => {
            File.WriteAllBytes(source, bytes);
            Assert.False(OfficeHeifMetadataReader.TryReadXmp(source, out text));
            Assert.Null(text);
            Assert.Equal(original, File.ReadAllBytes(source));
        });
        Assert.Equal(original, bytes);
    }

    [Fact]
    public void IndependentExifEditPreservesMalformedXmpPayloadBytes() {
        byte[] packet = { 0x61, 0xED, 0xA0, 0x80, 0x62 };
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "unused", 0, 0, xmpBytes: packet);
        var metadata = new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Software, "Changed");
        Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, metadata, out byte[]? edited));
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(edited!, out OfficeImageMetadata? reopened));
        Assert.Equal("Changed", reopened!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(edited!, out OfficeHeifImageInfo? info));
        OfficeHeifItemExtentInfo extent = info!.XmpItem!.Location!.Extents[0];
        Assert.Equal(packet, edited!.Skip(extent.Offset).Take(extent.Length).ToArray());
        Assert.False(OfficeHeifMetadataReader.TryReadXmp(edited!, out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnpairedUtf16XmpRejectsByteAndFileWritesWithoutMutation(bool highSurrogate) {
        string packet = "before" + (highSurrogate ? '\uD800' : '\uDC00') + "after";
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0);
        byte[] original = (byte[])bytes.Clone();
        Assert.False(OfficeHeifMetadataReader.TryWriteXmp(bytes, packet, out byte[]? output));
        Assert.Null(output);
        WithFiles((source, target) => {
            File.WriteAllBytes(source, bytes);
            byte[] sentinel = { 123, 45, 67 };
            File.WriteAllBytes(target, sentinel);
            Assert.False(OfficeHeifMetadataReader.TryWriteXmp(source, target, packet));
            Assert.Equal(sentinel, File.ReadAllBytes(target));
            Assert.False(OfficeHeifMetadataReader.TryWriteXmp(source, source, packet));
            Assert.Equal(original, File.ReadAllBytes(source));
        });
        Assert.Equal(original, bytes);
    }

    [Fact]
    public void ValidNonAsciiXmpRoundtripsAndCanReplaceMalformedRequestedPayload() {
        const string packet = "<xmp>Zażółć gęślą jaźń 日本語 \U0001F600</xmp>";
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "unused", 0, 0, xmpBytes: new byte[] { 0x80 });
        Assert.True(OfficeHeifMetadataReader.TryWriteXmp(bytes, packet, out byte[]? output));
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(output!, out string? reopened));
        Assert.Equal(packet, reopened);
        WithFiles((source, _) => {
            File.WriteAllBytes(source, output!);
            Assert.True(OfficeHeifMetadataReader.TryReadXmp(source, out reopened));
            Assert.Equal(packet, reopened);
        });
    }
}
