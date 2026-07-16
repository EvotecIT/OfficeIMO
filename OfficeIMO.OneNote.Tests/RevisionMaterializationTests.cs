namespace OfficeIMO.OneNote.Tests;

public sealed class RevisionMaterializationTests {
    [Fact]
    public void LaterRoleDeclarationSelectsIncrementalRevisionForActiveView() {
        OneNoteExtendedGuid objectSpaceId = Id(1);
        OneNoteRevisionManifest original = Revision(10, objectSpaceId, null, 1, 0);
        OneNoteRevisionManifest imported = Revision(11, objectSpaceId, original.Id, 1, 1);
        OneNoteRevisionManifest edited = Revision(12, objectSpaceId, imported.Id, 4, 2);
        edited.AddRoleAssociation(null, 1, 3);

        var emptyRoot = new OneNoteFileNodeList(0, Array.Empty<OneNoteFileNodeListFragment>(), Array.Empty<OneNoteFileNode>());
        var store = new OneNoteRevisionStore(
            new OneNoteFileHeader(),
            emptyRoot,
            new[] { emptyRoot },
            new[] { original, imported, edited },
            Array.Empty<OneNoteRevisionStoreObject>(),
            Array.Empty<OneNoteFileDataStoreObject>());

        OneNoteMaterializedObjectSpace? current = new OneNoteObjectSpaceMaterializer(store).TryGetCurrentSpace(objectSpaceId);

        Assert.NotNull(current);
        Assert.Same(edited, current.Revision);
        Assert.Equal(new[] { original.Id, imported.Id, edited.Id },
            new OneNoteObjectSpaceMaterializer(store).GetRevisionChain(edited).Select(item => item.Id));
    }

    [Fact]
    public void ContextRoleDoesNotReplaceDefaultContextActiveView() {
        OneNoteExtendedGuid objectSpaceId = Id(20);
        OneNoteExtendedGuid contextId = Id(21);
        OneNoteRevisionManifest active = Revision(22, objectSpaceId, null, 1, 0);
        OneNoteRevisionManifest history = Revision(23, objectSpaceId, active.Id, 1, 1, contextId);

        var emptyRoot = new OneNoteFileNodeList(0, Array.Empty<OneNoteFileNodeListFragment>(), Array.Empty<OneNoteFileNode>());
        var store = new OneNoteRevisionStore(
            new OneNoteFileHeader(),
            emptyRoot,
            new[] { emptyRoot },
            new[] { active, history },
            Array.Empty<OneNoteRevisionStoreObject>(),
            Array.Empty<OneNoteFileDataStoreObject>());
        var materializer = new OneNoteObjectSpaceMaterializer(store);

        Assert.Same(active, materializer.TryGetCurrentSpace(objectSpaceId)!.Revision);
        Assert.Same(history, materializer.TryGetSpace(objectSpaceId, contextId)!.Revision);
    }

    [Fact]
    public void ObjectReaderTraversesLongFileNodeListChainsWithoutCallStackRecursion() {
        const int depth = 20_000;
        var list = new OneNoteFileNodeList(
            depth,
            Array.Empty<OneNoteFileNodeListFragment>(),
            Array.Empty<OneNoteFileNode>());
        for (int index = depth - 1; index >= 0; index--) {
            var reference = new OneNoteFileNode(
                (ushort)OneNoteFileNodeId.ObjectSpaceManifestListReference,
                4,
                0,
                0,
                OneNoteFileNodeBaseType.FileNodeListReference,
                index,
                null,
                Array.Empty<byte>()) {
                ReferencedFileNodeList = list
            };
            list = new OneNoteFileNodeList(
                (uint)index,
                Array.Empty<OneNoteFileNodeListFragment>(),
                new[] { reference });
        }

        OneNoteRevisionStoreObjectReadResult result = OneNoteRevisionStoreObjectReader.Read(
            new MemoryStream(),
            list,
            0,
            new OneNoteReaderOptions());

        Assert.Empty(result.Revisions);
        Assert.Empty(result.Objects);
        Assert.Empty(result.FileDataObjects);
    }

    private static OneNoteRevisionManifest Revision(
        uint value,
        OneNoteExtendedGuid objectSpaceId,
        OneNoteExtendedGuid? dependencyId,
        uint role,
        int sourceOrder,
        OneNoteExtendedGuid? contextId = null) {
        var revision = new OneNoteRevisionManifest(Id(value)) {
            ObjectSpaceId = objectSpaceId,
            DependencyId = dependencyId,
            Role = role,
            ContextId = contextId
        };
        revision.AddRoleAssociation(contextId, role, sourceOrder);
        return revision;
    }

    private static OneNoteExtendedGuid Id(uint value) => new OneNoteExtendedGuid(
        new Guid("11111111-2222-3333-4444-555555555555"), value, 20);
}
