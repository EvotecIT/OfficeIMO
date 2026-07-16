namespace OfficeIMO.OneNote.Tests;

public sealed class PropertySetWriterTests {
    [Fact]
    public void DesktopWriterRoundTripsContextReferencesAndNestedPropertySets() {
        OneNoteExtendedGuid objectId = Id(1, "11111111-1111-1111-1111-111111111111");
        OneNoteExtendedGuid objectSpaceId = Id(2, "22222222-2222-2222-2222-222222222222");
        OneNoteExtendedGuid contextId = Id(3, "33333333-3333-3333-3333-333333333333");
        uint childElementId = Raw(OneNotePropertyType.PropertySet, 0x55);
        var nested = new[] {
            new OneNoteWriteProperty(Raw(OneNotePropertyType.ContextId, 4), references: new[] { contextId }, referenceKind: OneNoteWriteReferenceKind.Context, preserveRawId: true),
            new OneNoteWriteProperty(Raw(OneNotePropertyType.UInt32, 5), scalar: 42, preserveRawId: true)
        };
        var properties = new[] {
            new OneNoteWriteProperty(Raw(OneNotePropertyType.ObjectIdArray, 1), references: new[] { objectId }, preserveRawId: true),
            new OneNoteWriteProperty(Raw(OneNotePropertyType.ObjectSpaceId, 2), references: new[] { objectSpaceId }, referenceKind: OneNoteWriteReferenceKind.ObjectSpace, preserveRawId: true),
            new OneNoteWriteProperty(
                Raw(OneNotePropertyType.PropertySetArray, 3),
                childPropertySets: new[] { nested },
                childPropertyId: childElementId,
                preserveRawId: true)
        };
        var globalIds = new Dictionary<Guid, uint> {
            [objectId.Identifier] = 0,
            [objectSpaceId.Identifier] = 1,
            [contextId.Identifier] = 2
        };

        byte[] data = OneNotePropertySetWriter.WriteDesktop(properties, globalIds);
        OneNotePropertySet roundTrip = OneNotePropertySetReader.Read(
            data,
            globalIds.ToDictionary(item => item.Value, item => item.Key),
            new OneNoteReaderOptions(),
            0);

        Assert.Equal(objectId, Assert.Single(roundTrip.Properties[0].ReferencedIds));
        Assert.Equal(objectSpaceId, Assert.Single(roundTrip.Properties[1].ReferencedIds));
        OneNotePropertyValue nestedArray = roundTrip.Properties[2];
        Assert.Equal(childElementId, nestedArray.ChildPropertyId);
        OneNotePropertySet child = Assert.Single(nestedArray.ChildPropertySets);
        Assert.Equal(contextId, Assert.Single(child.Properties[0].ReferencedIds));
        Assert.Equal(42UL, child.Properties[1].ScalarValue);
    }

    private static uint Raw(OneNotePropertyType type, uint id) => ((uint)type << 26) | (id & 0x03FFFFFFU);

    private static OneNoteExtendedGuid Id(uint value, string guid) => new OneNoteExtendedGuid(new Guid(guid), value, 20);
}
