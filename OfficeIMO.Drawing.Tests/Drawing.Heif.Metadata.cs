using System;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingHeifMetadataTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(3)]
    public void ProfilesReadFromFileIdatAndMultipleExtents(int storage) {
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "<xmp>Original</xmp>", storage, storage);
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out OfficeImageMetadata? exif));
        Assert.Equal("Original", exif!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(bytes, out string? xmp));
        Assert.Equal("<xmp>Original</xmp>", xmp);
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.True(info!.HasExif && info.HasXmp);
        Assert.Equal(storage == 0, info.ExifItem!.Location!.CanWriteSingleFileExtent);
        Assert.Equal(storage == 3 ? 2 : 1, info.ExifItem.Location.Extents.Count);
    }

    [Fact]
    public void InfoAndStreamReadersPreservePrimaryPropertiesAndCallerPosition() {
        byte[] bytes = CoreHeifFixtures.CreateMinimalHeifWithPrimaryImageTransformProperties(
            320, 180, CoreHeifFixtures.CreateExifPayload("Original"));
        using var stream = new MemoryStream();
        stream.Write(new byte[] { 11, 22, 33 }, 0, 3);
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 3;
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(stream, out OfficeHeifImageInfo? info));
        Assert.Equal(320U, info!.Width);
        Assert.Equal(180U, info.Height);
        Assert.NotNull(info.RotationDegrees);
        Assert.True(info.IsMirrored);
        Assert.NotEmpty(info.PixelBitDepths);
        Assert.Equal(3L, stream.Position);
        Assert.True(stream.CanRead);
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(stream, out OfficeImageMetadata? profile));
        Assert.Equal("Original", profile!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.Equal(3L, stream.Position);
    }

    [Fact]
    public void DeclaredUnlocatedItemsRemainDiscoverableButUnreadable() {
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "packet", 2, 2);
        Assert.True(OfficeHeifMetadataReader.HasExifItem(bytes));
        Assert.True(OfficeHeifMetadataReader.HasXmpItem(bytes));
        Assert.False(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out _));
        Assert.False(OfficeHeifMetadataReader.TryReadXmp(bytes, out _));
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Null(info!.ExifItem!.Location);
        Assert.Null(info.XmpItem!.Location);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void RequestedWritesPreserveReadOnlyOrUnlocatedSibling(int siblingStorage) {
        byte[] exifSource = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "sibling-XMP", 0, siblingStorage);
        byte[] snapshot = (byte[])exifSource.Clone();
        var changed = new OfficeImageMetadata();
        changed.SetExifValue(OfficeExifTag.Software, "Changed");
        Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(exifSource, changed, out byte[]? exifOutput));
        Assert.Equal(snapshot, exifSource);
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(exifOutput!, out OfficeImageMetadata? exif));
        Assert.Equal("Changed", exif!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.Equal(siblingStorage != 2, OfficeHeifMetadataReader.TryReadXmp(exifOutput!, out string? siblingXmp));
        if (siblingStorage != 2) {
            Assert.Equal("sibling-XMP", siblingXmp);
        }

        byte[] xmpSource = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("sibling-EXIF"), "original-XMP", siblingStorage, 0);
        snapshot = (byte[])xmpSource.Clone();
        Assert.True(OfficeHeifMetadataReader.TryWriteXmp(xmpSource, "changed-XMP", out byte[]? xmpOutput));
        Assert.Equal(snapshot, xmpSource);
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(xmpOutput!, out string? xmp));
        Assert.Equal("changed-XMP", xmp);
        Assert.Equal(siblingStorage != 2, OfficeHeifMetadataReader.TryReadExifProfile(xmpOutput!, out OfficeImageMetadata? siblingExif));
        if (siblingStorage != 2) {
            Assert.Equal("sibling-EXIF", siblingExif!.GetExifValue(OfficeExifTag.Software)!.Value);
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void UnsupportedWriteLeavesInputAndExistingFileUntouched(int storage) {
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", storage, storage);
        byte[] original = (byte[])bytes.Clone();
        Assert.False(OfficeHeifMetadataReader.TryWriteXmp(bytes, "changed", out byte[]? output));
        Assert.Null(output);
        Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, null, out output));
        Assert.Null(output);
        Assert.Equal(original, bytes);
        WithFiles((source, target) => {
            File.WriteAllBytes(source, bytes);
            byte[] sentinel = { 123, 45, 67 };
            File.WriteAllBytes(target, sentinel);
            Assert.False(OfficeHeifMetadataReader.TryWriteXmp(source, target, "changed"));
            Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(source, target, null));
            Assert.Equal(sentinel, File.ReadAllBytes(target));
        });
    }

    [Fact]
    public void FileRoundtripReplacesAndClearsExistingProfiles() {
        WithFiles((source, target) => {
            File.WriteAllBytes(source, CoreHeifFixtures.CreateHeifMetadataSiblings(
                CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0));
            var metadata = new OfficeImageMetadata();
            metadata.SetExifValue(OfficeExifTag.Software, "Changed");
            Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(source, target, metadata));
            Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(target, out OfficeImageMetadata? reopened));
            Assert.Equal("Changed", reopened!.GetExifValue(OfficeExifTag.Software)!.Value);
            Assert.True(OfficeHeifMetadataReader.TryWriteXmp(target, target, "changed-XMP"));
            Assert.True(OfficeHeifMetadataReader.TryReadXmp(target, out string? xmp));
            Assert.Equal("changed-XMP", xmp);
            Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(target, target, null));
            Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(target, out reopened));
            Assert.Null(reopened);
            Assert.True(OfficeHeifMetadataReader.TryWriteXmp(target, target, null));
            Assert.True(OfficeHeifMetadataReader.TryReadXmp(target, out xmp));
            Assert.Equal(string.Empty, xmp);
        });
    }

    [Fact]
    public void CancelledPreparationLeavesExistingOutputUntouched() {
        WithFiles((source, target) => {
            File.WriteAllBytes(source, CoreHeifFixtures.CreateHeifMetadataSiblings(
                CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0));
            byte[] sentinel = { 123, 45, 67 };
            File.WriteAllBytes(target, sentinel);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => OfficeHeifMetadataReader.TryWriteXmp(
                source, target, "changed", cancellation.Token));
            Assert.Throws<OperationCanceledException>(() => OfficeHeifMetadataReader.TryReadInfo(
                File.ReadAllBytes(source), out _, cancellation.Token));
            Assert.Equal(sentinel, File.ReadAllBytes(target));
        });
    }

    [Fact]
    public void ExcessiveDeclaredItemCountFailsBeforeBuildingCollections() {
        byte[] bytes = CoreHeifFixtures.CreateMinimalHeifWithExif(CoreHeifFixtures.CreateExifPayload("Original"));
        byte[] marker = System.Text.Encoding.ASCII.GetBytes("iinf");
        int at = Enumerable.Range(0, bytes.Length - marker.Length).First(index =>
            bytes.Skip(index).Take(marker.Length).SequenceEqual(marker));
        bytes[at + 8] = 0x13;
        bytes[at + 9] = 0x88; // 5,000 declared items, above the bounded collection contract.
        Assert.False(OfficeHeifMetadataReader.TryReadInfo(bytes, out _));
        Assert.False(OfficeHeifMetadataReader.HasExifItem(bytes));
        Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, null, out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedExtentsCannotEraseSiblingMetadataOrImagePixels(bool imageSibling) {
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0);
        int locationType = FindAscii(bytes, "iloc");
        Buffer.BlockCopy(bytes, locationType + 20, bytes, locationType + 36, 8);
        if (imageSibling) {
            Buffer.BlockCopy(System.Text.Encoding.ASCII.GetBytes("hvc1"), 0, bytes, FindAscii(bytes, "mime"), 4);
        }
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Equal(imageSibling ? "hvc1" : "mime", info!.Items[1].ItemType);
        Assert.Equal(info.Items[0].Location!.Extents[0].Offset, info.Items[1].Location!.Extents[0].Offset);
        byte[] original = (byte[])bytes.Clone();
        var metadata = new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Software, "Changed");
        Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, metadata, out byte[]? output));
        Assert.Null(output);
        Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, null, out output));
        Assert.Null(output);
        Assert.Equal(original, bytes);
        WithFiles((source, target) => {
            File.WriteAllBytes(source, bytes);
            byte[] sentinel = { 123, 45, 67 };
            File.WriteAllBytes(target, sentinel);
            Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(source, target, metadata));
            Assert.Equal(sentinel, File.ReadAllBytes(target));
        });
    }

    private static int FindAscii(byte[] data, string text) {
        byte[] marker = System.Text.Encoding.ASCII.GetBytes(text);
        return Enumerable.Range(0, data.Length - marker.Length).First(index =>
            data.Skip(index).Take(marker.Length).SequenceEqual(marker));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReplacementPayloadUsesFramedMediaDataBox(bool zeroSizedTail) {
        byte[] bytes = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0);
        if (zeroSizedTail) {
            Array.Clear(bytes, FindAscii(bytes, "mdat") - 4, 4);
        }
        var metadata = new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Software, "Changed");
        Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, metadata, out byte[]? output));
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(output!, out OfficeHeifImageInfo? info));
        Assert.Equal(bytes.Length + 8, info!.ExifItem!.Location!.Extents[0].Offset);
        AssertFramedMediaData(output!, bytes.Length);
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(output!, out OfficeImageMetadata? exif));
        Assert.Equal("Changed", exif!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(output!, out string? sibling));
        Assert.Equal("original-XMP", sibling);
        Assert.True(OfficeHeifMetadataReader.TryWriteXmp(output!, "changed-XMP", out byte[]? next));
        AssertFramedMediaData(next!, output!.Length);
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(next!, out string? xmp));
        Assert.Equal("changed-XMP", xmp);
    }

    private static void AssertFramedMediaData(byte[] data, int appendedBoxOffset) {
        // Independent BMFF framing walk; all boxes consume the output, with no raw tail.
        using var stream = new MemoryStream(data, writable: false);
        using var reader = new BinaryReader(stream);
        while (stream.Position < stream.Length) {
            long offset = stream.Position;
            byte[] sizeBytes = reader.ReadBytes(4);
            Assert.Equal(4, sizeBytes.Length);
            uint size = ((uint)sizeBytes[0] << 24) | ((uint)sizeBytes[1] << 16) |
                ((uint)sizeBytes[2] << 8) | sizeBytes[3];
            string type = System.Text.Encoding.ASCII.GetString(reader.ReadBytes(4));
            Assert.True(size >= 8 && size <= stream.Length - offset);
            if (offset == appendedBoxOffset) {
                Assert.Equal("mdat", type);
                Assert.Equal(stream.Length, offset + size);
            }
            stream.Position = offset + size;
        }
        Assert.Equal(stream.Length, stream.Position);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AppendingPastEncodedLimitRejectsBeforeReturningOrWritingOutput(bool fileApi) {
        byte[] fixture = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "original-XMP", 0, 0);
        if (!fileApi) {
            var nearLimit = new byte[OfficeRasterGuards.MaximumEncodedBytes];
            Buffer.BlockCopy(fixture, 0, nearLimit, 0, fixture.Length);
            Assert.False(OfficeHeifMetadataReader.TryWriteXmp(nearLimit, "changed", out byte[]? output));
            Assert.Null(output);
            Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(nearLimit,
                OfficeImageMetadata.ParseExifProfile(CoreHeifFixtures.CreateExifPayload("Changed")), out output));
            Assert.Null(output);
            Assert.Equal(fixture, nearLimit.Take(fixture.Length).ToArray());
            return;
        }
        WithFiles((source, target) => {
            using (FileStream stream = File.Create(source)) {
                stream.Write(fixture, 0, fixture.Length);
                stream.SetLength(OfficeRasterGuards.MaximumEncodedBytes);
            }
            byte[] sentinel = { 123, 45, 67 };
            File.WriteAllBytes(target, sentinel);
            Assert.False(OfficeHeifMetadataReader.TryWriteXmp(source, target, "changed"));
            Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(source, target,
                OfficeImageMetadata.ParseExifProfile(CoreHeifFixtures.CreateExifPayload("Changed"))));
            Assert.Equal(sentinel, File.ReadAllBytes(target));
        });
    }

    private static void WithFiles(Action<string, string> action) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-core-heif-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            action(Path.Combine(directory, "source.heic"), Path.Combine(directory, "output.heic"));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
