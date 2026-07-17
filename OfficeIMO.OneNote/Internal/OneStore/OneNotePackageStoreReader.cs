namespace OfficeIMO.OneNote;

internal static partial class OneNotePackageStoreReader {
    private const int DataElement = 0x01;
    private const int DataElementPackage = 0x15;
    private const ulong StorageIndexType = 0x01;
    private const ulong CellManifestType = 0x03;
    private const ulong RevisionManifestType = 0x04;
    private const ulong ObjectGroupType = 0x05;
    private const ulong ObjectBlobType = 0x0A;

    internal static OneNoteRevisionStore Read(Stream stream, OneNoteFileHeader header, OneNoteReaderOptions options) {
        FssHttpStreamObject packaging = FssHttpStreamObjectReader.ReadPackaging(stream, options);
        FssHttpStreamObject package = RequireExactlyOne(
            packaging.Children,
            item => item.Type == DataElementPackage,
            "ONENOTE_PACKAGE_DATA_ELEMENTS",
            "The package does not contain exactly one data-element package.",
            packaging.HeaderOffset);
        List<PackageDataElement> elements = ReadDataElements(stream, package, options);
        PackageDataElement storageIndex = RequireExactlyOne(
            elements,
            item => item.Type == StorageIndexType,
            "ONENOTE_PACKAGE_STORAGE_INDEX",
            "The package does not contain exactly one storage-index data element.",
            package.HeaderOffset);
        PackageIndex index = ReadStorageIndex(stream, storageIndex, options);
        PackageGraph graph = ReconstructGraph(stream, elements, index, options);

        var emptyRoot = new OneNoteFileNodeList(0, Array.Empty<OneNoteFileNodeListFragment>(), Array.Empty<OneNoteFileNode>());
        return new OneNoteRevisionStore(
            header,
            emptyRoot,
            new[] { emptyRoot },
            graph.Revisions.AsReadOnly(),
            graph.Objects.AsReadOnly(),
            graph.FileDataObjects.AsReadOnly());
    }

    private static List<PackageDataElement> ReadDataElements(Stream stream, FssHttpStreamObject package, OneNoteReaderOptions options) {
        var elements = new List<PackageDataElement>();
        var ids = new HashSet<PackageGuidKey>();
        foreach (FssHttpStreamObject item in package.Children) {
            if (item.Type != DataElement || !item.Compound) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_DATA_ELEMENT", "The data-element package contains an unexpected stream object.", item.HeaderOffset);
            }
            byte[] prefix = FssHttpStreamObjectReader.ReadData(stream, item, 128, "data-element prefix");
            var cursor = new FssHttpDataCursor(prefix, item.DataOffset);
            OneNoteExtendedGuid id = cursor.ReadExtendedGuid();
            cursor.SkipSerialNumber();
            ulong type = cursor.ReadCompactUInt64();
            cursor.EnsureEnd("data-element prefix");
            if (id.Identifier == Guid.Empty || !ids.Add(new PackageGuidKey(id))) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_DATA_ELEMENT_ID", "A data-element identifier is empty or duplicated.", item.DataOffset);
            }
            elements.Add(new PackageDataElement(id, type, item));
        }
        if (elements.Count > options.MaxStreamObjects) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_DATA_ELEMENT_LIMIT", "The data-element count exceeds the configured stream-object limit.", package.HeaderOffset);
        }
        return elements;
    }

    private static PackageGraph ReconstructGraph(Stream stream, List<PackageDataElement> elements, PackageIndex index, OneNoteReaderOptions options) {
        var byElementId = elements.ToDictionary(item => new PackageGuidKey(item.Id));
        var revisionElements = elements.Where(item => item.Type == RevisionManifestType)
            .Select(item => ReadRevisionManifest(stream, item, options))
            .ToList();
        var revisionByElementId = revisionElements.ToDictionary(item => new PackageGuidKey(item.Element.Id));
        var revisionById = new Dictionary<PackageGuidKey, PackageRevisionElement>();
        foreach (PackageRevisionElement revision in revisionElements) {
            var key = new PackageGuidKey(revision.Manifest.Id);
            if (revisionById.ContainsKey(key)) {
                throw new OneNoteFormatException(
                    "ONENOTE_PACKAGE_REVISION_ID",
                    "A package revision-manifest identifier is duplicated.",
                    revision.Element.Node.HeaderOffset);
            }
            revisionById.Add(key, revision);
        }

        int roleAssociationOrder = 0;
        foreach (PackageCellMapping cellMapping in index.CellMappings) {
            if (!byElementId.TryGetValue(new PackageGuidKey(cellMapping.CellManifestElementId), out PackageDataElement? cellElement) || cellElement.Type != CellManifestType) continue;
            OneNoteExtendedGuid currentRevisionMappingId = ReadCellManifest(stream, cellElement, options);
            if (!index.RevisionMappings.TryGetValue(new PackageGuidKey(currentRevisionMappingId), out OneNoteExtendedGuid? revisionElementId)) continue;
            if (!revisionByElementId.TryGetValue(new PackageGuidKey(revisionElementId), out PackageRevisionElement? current)) continue;
            AssignRevisionChain(current, cellMapping.CellId, revisionById, roleAssociationOrder++);
        }

        var graph = new PackageGraph();
        foreach (PackageRevisionElement revision in revisionElements.Where(item => item.Manifest.ObjectSpaceId != null)) {
            graph.Revisions.Add(revision.Manifest);
        }

        var objectGroups = elements.Where(item => item.Type == ObjectGroupType)
            .ToDictionary(item => new PackageGuidKey(item.Id));
        var blobs = elements.Where(item => item.Type == ObjectBlobType)
            .ToDictionary(item => new PackageGuidKey(item.Id));
        var seenObjects = new HashSet<PackageGuidKey>();
        foreach (PackageRevisionElement revision in revisionElements.Where(item => item.Manifest.ObjectSpaceId != null)) {
            foreach (OneNoteExtendedGuid groupId in revision.ObjectGroupIds) {
                if (!objectGroups.TryGetValue(new PackageGuidKey(groupId), out PackageDataElement? group)) continue;
                ReadObjectGroup(stream, group, revision, blobs, graph, seenObjects, options);
            }
        }
        return graph;
    }

    private static void AssignRevisionChain(
        PackageRevisionElement current,
        FssHttpCellId cellId,
        IReadOnlyDictionary<PackageGuidKey, PackageRevisionElement> revisionsById,
        int roleAssociationOrder) {
        var visited = new HashSet<PackageGuidKey>();
        PackageRevisionElement? revision = current;
        bool isCurrent = true;
        while (revision != null && visited.Add(new PackageGuidKey(revision.Manifest.Id))) {
            revision.CellId = cellId;
            revision.Manifest.ObjectSpaceId = cellId.Second;
            revision.Manifest.ContextId = IsDefaultContext(cellId.First) ? null : cellId.First;
            if (isCurrent) {
                revision.Manifest.Role = 1;
                revision.Manifest.AddRoleAssociation(revision.Manifest.ContextId, 1, roleAssociationOrder);
                isCurrent = false;
            }
            revision = revision.Manifest.DependencyId != null && revisionsById.TryGetValue(new PackageGuidKey(revision.Manifest.DependencyId), out PackageRevisionElement? dependency)
                ? dependency
                : null;
        }
    }

    private static bool IsDefaultContext(OneNoteExtendedGuid value) {
        return (value.Identifier == Guid.Empty && value.Value == 0) ||
               (value.Identifier == new Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073") && value.Value == 1);
    }

    private static OneNoteExtendedGuid ReadCellManifest(Stream stream, PackageDataElement element, OneNoteReaderOptions options) {
        FssHttpStreamObject current = RequireExactlyOne(
            element.Node.Children,
            item => item.Type == 0x0B,
            "ONENOTE_PACKAGE_CELL_MANIFEST",
            "A cell manifest must declare exactly one current revision.",
            element.Node.HeaderOffset);
        byte[] data = FssHttpStreamObjectReader.ReadData(stream, current, 32, "cell-manifest current revision");
        var cursor = new FssHttpDataCursor(data, current.DataOffset);
        OneNoteExtendedGuid revision = cursor.ReadExtendedGuid();
        cursor.EnsureEnd("cell-manifest current revision");
        return revision;
    }

    private static PackageRevisionElement ReadRevisionManifest(Stream stream, PackageDataElement element, OneNoteReaderOptions options) {
        FssHttpStreamObject revisionNode = RequireExactlyOne(
            element.Node.Children,
            item => item.Type == 0x1A,
            "ONENOTE_PACKAGE_REVISION_MANIFEST",
            "A revision-manifest data element must contain exactly one revision declaration.",
            element.Node.HeaderOffset);
        byte[] data = FssHttpStreamObjectReader.ReadData(stream, revisionNode, 64, "revision manifest");
        var cursor = new FssHttpDataCursor(data, revisionNode.DataOffset);
        OneNoteExtendedGuid revisionId = cursor.ReadExtendedGuid();
        OneNoteExtendedGuid dependency = cursor.ReadExtendedGuid();
        cursor.EnsureEnd("revision manifest");
        var manifest = new OneNoteRevisionManifest(revisionId) {
            DependencyId = IsNull(dependency) ? null : dependency
        };
        var result = new PackageRevisionElement(element, manifest);
        foreach (FssHttpStreamObject root in element.Node.Children.Where(item => item.Type == 0x0A)) {
            byte[] rootData = FssHttpStreamObjectReader.ReadData(stream, root, 64, "revision root declaration");
            var rootCursor = new FssHttpDataCursor(rootData, root.DataOffset);
            OneNoteExtendedGuid rootId = rootCursor.ReadExtendedGuid();
            OneNoteExtendedGuid objectId = rootCursor.ReadExtendedGuid();
            rootCursor.EnsureEnd("revision root declaration");
            manifest.RootObjects.Add(new OneNoteRootObjectReference(objectId, rootId.Value));
        }
        foreach (FssHttpStreamObject group in element.Node.Children.Where(item => item.Type == 0x19)) {
            byte[] groupData = FssHttpStreamObjectReader.ReadData(stream, group, 32, "revision object-group reference");
            var groupCursor = new FssHttpDataCursor(groupData, group.DataOffset);
            result.ObjectGroupIds.Add(groupCursor.ReadExtendedGuid());
            groupCursor.EnsureEnd("revision object-group reference");
        }
        return result;
    }

    private static bool IsNull(OneNoteExtendedGuid value) => value.Identifier == Guid.Empty && value.Value == 0;

    private static T RequireExactlyOne<T>(
        IEnumerable<T> source,
        Func<T, bool> predicate,
        string code,
        string message,
        long offset) {
        using IEnumerator<T> matches = source.Where(predicate).GetEnumerator();
        if (!matches.MoveNext()) throw new OneNoteFormatException(code, message, offset);
        T value = matches.Current;
        if (matches.MoveNext()) throw new OneNoteFormatException(code, message, offset);
        return value;
    }

    private sealed class PackageDataElement {
        internal PackageDataElement(OneNoteExtendedGuid id, ulong type, FssHttpStreamObject node) { Id = id; Type = type; Node = node; }
        internal OneNoteExtendedGuid Id { get; }
        internal ulong Type { get; }
        internal FssHttpStreamObject Node { get; }
    }

    private sealed class PackageRevisionElement {
        internal PackageRevisionElement(PackageDataElement element, OneNoteRevisionManifest manifest) { Element = element; Manifest = manifest; }
        internal PackageDataElement Element { get; }
        internal OneNoteRevisionManifest Manifest { get; }
        internal FssHttpCellId? CellId { get; set; }
        internal List<OneNoteExtendedGuid> ObjectGroupIds { get; } = new List<OneNoteExtendedGuid>();
    }

    private sealed class PackageGraph {
        internal List<OneNoteRevisionManifest> Revisions { get; } = new List<OneNoteRevisionManifest>();
        internal List<OneNoteRevisionStoreObject> Objects { get; } = new List<OneNoteRevisionStoreObject>();
        internal List<OneNoteFileDataStoreObject> FileDataObjects { get; } = new List<OneNoteFileDataStoreObject>();
        internal long TotalAssetBytes { get; set; }
    }

    private readonly struct PackageGuidKey : IEquatable<PackageGuidKey> {
        internal PackageGuidKey(OneNoteExtendedGuid value) { Identifier = value.Identifier; Value = value.Value; }
        private Guid Identifier { get; }
        private uint Value { get; }
        public bool Equals(PackageGuidKey other) => Identifier == other.Identifier && Value == other.Value;
        public override bool Equals(object? obj) => obj is PackageGuidKey other && Equals(other);
        public override int GetHashCode() => (Identifier.GetHashCode() * 397) ^ Value.GetHashCode();
    }
}
