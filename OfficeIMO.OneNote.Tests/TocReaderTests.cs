using System.Text;

namespace OfficeIMO.OneNote.Tests;

public sealed class TocReaderTests {
    [Fact]
    public void MapsOrderedSectionAndSectionGroupEntries() {
        OneNoteExtendedGuid rootId = Extended(1);
        OneNoteExtendedGuid sectionId = Extended(2);
        OneNoteExtendedGuid groupId = Extended(3);
        OneNoteRevisionStoreObject root = Object(
            rootId,
            Property(0x24001CF6, referencedIds: new[] { groupId, sectionId }),
            Property(0x14001CBE, scalarValue: 0x0097982EU),
            Property(0x88001E1E, booleanValue: true));
        Guid expectedSectionId = Guid.NewGuid();
        OneNoteRevisionStoreObject section = Object(
            sectionId,
            Property(0x1C001D94, data: expectedSectionId.ToByteArray()),
            Property(0x14001CB9, scalarValue: 1),
            Property(0x1C001D6B, data: Unicode("Planning.one")),
            Property(0x14001CBE, scalarValue: 0x00123456));
        OneNoteRevisionStoreObject group = Object(
            groupId,
            Property(0x1C001D94, data: Guid.NewGuid().ToByteArray()),
            Property(0x14001CB9, scalarValue: 2),
            Property(0x1C001D6B, data: Unicode("Archive")),
            Property(0x14001CBE, scalarValue: 0xFFFFFFFF));

        OneNoteTocData toc = OneNoteTocMapper.Map(Store(root, section, group));

        Assert.Equal(0x0097982EU, toc.ColorArgb);
        Assert.True(toc.HistoryEnabled);
        Assert.Collection(toc.Entries,
            item => {
                Assert.Equal("Archive", item.Name);
                Assert.False(item.IsSection);
                Assert.Null(item.ColorArgb);
            },
            item => {
                Assert.Equal("Planning.one", item.Name);
                Assert.True(item.IsSection);
                Assert.Equal(expectedSectionId, item.Id);
                Assert.Equal(1U, item.Order);
                Assert.Equal(0x00123456U, item.ColorArgb);
            });
        Assert.Equal(3, toc.PreservedObjects.Count);
    }

    private static OneNoteRevisionStore Store(params OneNoteRevisionStoreObject[] objects) {
        OneNoteExtendedGuid objectSpaceId = Extended(20);
        OneNoteExtendedGuid revisionId = Extended(21);
        var revision = new OneNoteRevisionManifest(revisionId) {
            ObjectSpaceId = objectSpaceId,
            Role = 1
        };
        revision.AddRoleAssociation(null, 1, 0);
        revision.RootObjects.Add(new OneNoteRootObjectReference(objects[0].Id, 1));
        foreach (OneNoteRevisionStoreObject item in objects) item.RevisionId = revisionId;
        var rootList = new OneNoteFileNodeList(1, Array.Empty<OneNoteFileNodeListFragment>(), Array.Empty<OneNoteFileNode>());
        return new OneNoteRevisionStore(
            new OneNoteFileHeader { FileKind = OneNoteFileKind.TableOfContents },
            rootList,
            new[] { rootList },
            new[] { revision },
            objects,
            Array.Empty<OneNoteFileDataStoreObject>());
    }

    private static OneNoteRevisionStoreObject Object(OneNoteExtendedGuid id, params OneNotePropertyValue[] properties) {
        var node = new OneNoteFileNode(0, 4, 0, 0, OneNoteFileNodeBaseType.Inline, 0, null, Array.Empty<byte>());
        return new OneNoteRevisionStoreObject(id, new OneNoteJcid(0x00020001), node) {
            PropertySet = new OneNotePropertySet(properties, 0),
            RawPropertyData = OneNoteBinaryPayload.FromBytes(new byte[] { 0 })
        };
    }

    private static OneNotePropertyValue Property(
        uint rawId,
        bool? booleanValue = null,
        ulong? scalarValue = null,
        byte[]? data = null,
        IReadOnlyList<OneNoteExtendedGuid>? referencedIds = null) {
        var property = new OneNotePropertyValue(rawId, 0) {
            BooleanValue = booleanValue,
            ScalarValue = scalarValue,
            ReferencedIds = referencedIds ?? Array.Empty<OneNoteExtendedGuid>()
        };
        if (data != null) property.Data = OneNoteBinaryPayload.FromBytes(data);
        return property;
    }

    private static OneNoteExtendedGuid Extended(uint value) => new OneNoteExtendedGuid(Guid.Parse("00112233-4455-6677-8899-aabbccddeeff"), value, 4);
    private static byte[] Unicode(string value) => Encoding.Unicode.GetBytes(value + "\0");
}
