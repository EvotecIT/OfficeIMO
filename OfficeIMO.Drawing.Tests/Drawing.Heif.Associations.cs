using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingHeifMetadataTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MetadataAssociationsSelectPrimaryRegardlessOfDeclarationOrder(bool largeIds) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(largeIds);
        uint first = largeIds ? 70000U : 1U;
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out OfficeImageMetadata? exif));
        Assert.Equal("Primary", exif!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(bytes, out string? xmp));
        Assert.Equal("primary-XMP", xmp);
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? before));
        Assert.Equal(first + 2, before!.ExifItem!.ItemId);
        Assert.Equal(first + 4, before.XmpItem!.ItemId);
        Assert.Equal(6, before.Items.Count);
        var metadata = new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Software, "Changed");
        Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, metadata, out byte[]? edited));
        AssertUntargetedPayloads(bytes, edited!, before, first + 2);
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(edited!, out exif));
        Assert.Equal("Changed", exif!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.True(OfficeHeifMetadataReader.TryWriteXmp(bytes, "changed-XMP", out edited));
        AssertUntargetedPayloads(bytes, edited!, before, first + 4);
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(edited!, out xmp));
        Assert.Equal("changed-XMP", xmp);
        WithFiles((source, target) => {
            File.WriteAllBytes(source, bytes);
            Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(source, target, metadata));
            Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(target, out exif));
            Assert.Equal("Changed", exif!.GetExifValue(OfficeExifTag.Software)!.Value);
            Assert.Equal(bytes, File.ReadAllBytes(source));
        });
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(5)]
    public void MetadataAssociationsRejectAmbiguityWithoutHidingDeclarations(int associations) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(false, associations);
        Assert.True(OfficeHeifMetadataReader.HasExifItem(bytes));
        Assert.True(OfficeHeifMetadataReader.HasXmpItem(bytes));
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.True(info!.HasExif && info.HasXmp);
        Assert.Equal(6, info.Items.Count);
        Assert.Null(info.ExifItem);
        Assert.Null(info.XmpItem);
        AssertMetadataRejectionPreserves(bytes);
    }

    [Fact]
    public void MetadataAssociationsKeepUniqueLegacyItemsAndExcludeExplicitThumbnailItems() {
        byte[] unique = CoreHeifFixtures.CreateHeifAssociationGraph(false, 2, uniqueMetadata: true);
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(unique, out OfficeImageMetadata? exif));
        Assert.Equal("Primary", exif!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(unique, out string? xmp));
        Assert.Equal("primary-XMP", xmp);
        AssertMetadataRejectionPreserves(CoreHeifFixtures.CreateHeifAssociationGraph(false, 4, uniqueMetadata: true));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MetadataAssociationsDoNotFallBackFromUnsupportedPrimaryToReadableThumbnail(bool exifProtected) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(false,
            protectedPrimaryExif: exifProtected, encodedPrimaryXmp: !exifProtected);
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Equal(3U, info!.ExifItem!.ItemId);
        Assert.Equal(5U, info.XmpItem!.ItemId);
        if (exifProtected) {
            Assert.True(OfficeHeifMetadataReader.HasExifItem(bytes));
            Assert.False(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out _));
            Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, null, out _));
            Assert.True(OfficeHeifMetadataReader.TryWriteXmp(bytes, "changed", out byte[]? edited));
            AssertUntargetedPayloads(bytes, edited!, info, 5);
        } else {
            Assert.True(OfficeHeifMetadataReader.HasXmpItem(bytes));
            Assert.False(OfficeHeifMetadataReader.TryReadXmp(bytes, out _));
            Assert.False(OfficeHeifMetadataReader.TryWriteXmp(bytes, null, out _));
            Assert.True(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, null, out byte[]? edited));
            AssertUntargetedPayloads(bytes, edited!, info, 3);
        }
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(false, 3)]
    [InlineData(false, 4)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    [InlineData(true, 3)]
    [InlineData(true, 4)]
    public void ReferenceDeclarationsRejectTruncationInsteadOfPublishingPartialCollections(bool largeIds, int malformed) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(largeIds, malformedReference: malformed);
        Assert.False(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Null(info);
        AssertMetadataRejectionPreserves(bytes);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void IncompleteStructureCollectionsNeverPublishPartialInformation(int malformedCollection) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(false, malformedCollection: malformedCollection);
        Assert.False(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Null(info);
        if (malformedCollection != 3) {
            AssertMetadataRejectionPreserves(bytes);
        }
    }

    private static void AssertUntargetedPayloads(byte[] source, byte[] edited, OfficeHeifImageInfo before, uint editedItem) {
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(edited, out OfficeHeifImageInfo? after));
        Assert.Equal(before.Items.Select(item => item.ItemId), after!.Items.Select(item => item.ItemId));
        foreach (OfficeHeifItemInfo item in before.Items.Where(item => item.ItemId != editedItem)) {
            OfficeHeifItemExtentInfo original = item.Location!.Extents[0];
            OfficeHeifItemExtentInfo current = after.Items.Single(value => value.ItemId == item.ItemId).Location!.Extents[0];
            // Compare raw independently located payload bytes, including both opaque image packets.
            Assert.Equal(source.Skip(original.Offset).Take(original.Length).ToArray(), edited.Skip(current.Offset).Take(current.Length).ToArray());
        }
    }

    private static void AssertMetadataRejectionPreserves(byte[] bytes) {
        byte[] original = (byte[])bytes.Clone();
        Assert.False(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out _));
        Assert.False(OfficeHeifMetadataReader.TryReadXmp(bytes, out _));
        Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(bytes, null, out byte[]? output));
        Assert.Null(output);
        Assert.False(OfficeHeifMetadataReader.TryWriteXmp(bytes, "changed", out output));
        Assert.Null(output);
        Assert.Equal(original, bytes);
        WithFiles((source, target) => {
            File.WriteAllBytes(source, bytes);
            byte[] sentinel = { 123, 45, 67 };
            File.WriteAllBytes(target, sentinel);
            Assert.False(OfficeHeifMetadataReader.TryWriteXmp(source, target, "changed"));
            Assert.False(OfficeHeifMetadataReader.TryWriteExifProfile(source, target, null));
            Assert.Equal(sentinel, File.ReadAllBytes(target));
            Assert.Equal(original, File.ReadAllBytes(source));
        });
    }
}
