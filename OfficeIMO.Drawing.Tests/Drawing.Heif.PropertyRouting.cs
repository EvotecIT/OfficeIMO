using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingHeifMetadataTests {
    [Theory]
    [InlineData(false, 1, 0)]
    [InlineData(false, 2, 0)]
    [InlineData(false, 0, 1)]
    [InlineData(true, 1, 1)]
    [InlineData(true, 2, 1)]
    [InlineData(true, 0, 1)]
    [InlineData(false, 0, 3)]
    [InlineData(false, 1, 3)]
    [InlineData(true, 2, 3)]
    public void PropertyRoutingValidatesEveryListBeforeReturningInformation(bool largeIds, int containerMode, int associationMode) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(largeIds,
            malformedCollection: associationMode == 3 ? 0 : 3,
            propertyContainerMode: containerMode, propertyAssociationMode: associationMode);
        byte[] original = (byte[])bytes.Clone();
        Assert.False(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Null(info);
        using var stream = new MemoryStream(bytes);
        Assert.False(OfficeHeifMetadataReader.TryReadInfo(stream, out info));
        Assert.Null(info);
        Assert.Equal(0, stream.Position);
        Assert.True(stream.CanRead);
        WithFiles((source, target) => {
            File.WriteAllBytes(source, bytes);
            Assert.False(OfficeHeifMetadataReader.TryReadInfo(source, out info));
            Assert.Null(info);
            Assert.Equal(original, File.ReadAllBytes(source));
        });
        Assert.Equal(original, bytes);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void PropertyRoutingKeepsValidMetadataOnlyContainersWithoutPropertyStorage(bool largeIds, int containerMode) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(largeIds, propertyContainerMode: containerMode);
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Null(info!.Width);
        Assert.True(info.HasExif && info.HasXmp);
        Assert.Equal(6, info.Items.Count);
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out OfficeImageMetadata? profile));
        Assert.Equal("Primary", profile!.GetExifValue(OfficeExifTag.Software)!.Value);
        Assert.True(OfficeHeifMetadataReader.TryReadXmp(bytes, out string? packet));
        Assert.Equal("primary-XMP", packet);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PropertyRoutingRetainsDisjointItemsFromAllValidAssociationLists(bool largeIds) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(largeIds, propertyAssociationMode: 1);
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Equal(640U, info!.Width);
        uint thumbnail = largeIds ? 70001U : 2U;
        OfficeHeifItemInfo item = info.Items.Single(value => value.ItemId == thumbnail);
        Assert.Equal(640U, item.Width);
        Assert.Equal(480U, item.Height);
        Assert.Single(info.PrimaryItem!.PropertyAssociations);
        Assert.Single(item.PropertyAssociations);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PropertyRoutingKeepsFirstEntryForRepeatedItemsAcrossCompleteLists(bool largeIds) {
        byte[] bytes = CoreHeifFixtures.CreateHeifAssociationGraph(largeIds, propertyAssociationMode: 2);
        Assert.True(OfficeHeifMetadataReader.TryReadInfo(bytes, out OfficeHeifImageInfo? info));
        Assert.Equal(640U, info!.Width);
        Assert.Null(info.PrimaryItem!.RotationDegrees);
        Assert.Equal("ispe", Assert.Single(info.PrimaryItem.PropertyAssociations).PropertyType);
    }
}
