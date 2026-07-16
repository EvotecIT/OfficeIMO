using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.OneNote;

internal sealed class OneNoteDesktopWritePlan {
    private static readonly Guid FileDataHeader = new Guid("BDE316E7-2665-4511-A4C4-8D4D0B7A9EAC");
    private static readonly Guid FileDataFooter = new Guid("71FBA722-0F79-4A0B-BB13-899256426B24");
    private readonly OneNoteWriteGraph _graph;
    private readonly List<OneNoteDesktopFileNodeList> _lists = new List<OneNoteDesktopFileNodeList>();
    private readonly List<OneNoteDesktopDataChunk> _data = new List<OneNoteDesktopDataChunk>();
    private readonly OneNoteDesktopFileNodeList _root;
    private OneNoteDesktopFileNodeList? _fileDataList;
    private uint _nextListId = 0x10;

    private OneNoteDesktopWritePlan(OneNoteWriteGraph graph) {
        _graph = graph;
        _root = NewList();
    }

    internal OneNoteWriteGraph Graph => _graph;
    internal IReadOnlyList<OneNoteDesktopFileNodeList> Lists => _lists;
    internal IReadOnlyList<OneNoteDesktopDataChunk> Data => _data;
    internal OneNoteDesktopFileNodeList Root => _root;

    internal static OneNoteDesktopWritePlan Create(OneNoteWriteGraph graph) {
        if (graph == null) throw new ArgumentNullException(nameof(graph));
        if (graph.FileKind != OneNoteFileKind.Section && graph.FileKind != OneNoteFileKind.TableOfContents) {
            throw new OneNoteFormatException("ONENOTE_WRITE_FILE_KIND", "A desktop revision store can only represent a .one section or .onetoc2 table of contents.");
        }
        if (graph.ObjectSpaces.Count == 0) throw new OneNoteFormatException("ONENOTE_WRITE_OBJECT_SPACE", "The desktop OneNote graph contains no object spaces.");
        if (graph.FileKind == OneNoteFileKind.TableOfContents && graph.ObjectSpaces.Count != 1) {
            throw new OneNoteFormatException("ONENOTE_WRITE_TOC_OBJECT_SPACE", "A .onetoc2 write graph must contain exactly one object space.");
        }

        var plan = new OneNoteDesktopWritePlan(graph);
        OneNoteWriteObjectSpace[][] objectSpaceGroups = graph.ObjectSpaces
            .GroupBy(space => space.Id)
            .Select(group => group.ToArray())
            .ToArray();
        plan.AddObjectSpace(objectSpaceGroups[0]);
        plan._root.Nodes.Add(OneNoteDesktopFileNode.Inline(
            OneNoteFileNodeId.ObjectSpaceManifestRoot,
            OneNoteDesktopBinary.Data(stream => OneNoteDesktopBinary.WriteExtendedGuid(stream, graph.RootObjectSpaceId))));
        for (int index = 1; index < objectSpaceGroups.Length; index++) plan.AddObjectSpace(objectSpaceGroups[index]);
        if (plan._fileDataList != null && plan._fileDataList.Nodes.Count > 0) {
            plan._root.Nodes.Add(OneNoteDesktopFileNode.List(OneNoteFileNodeId.FileDataStoreListReference, plan._fileDataList));
        }
        return plan;
    }

    private void AddObjectSpace(IReadOnlyList<OneNoteWriteObjectSpace> spaces) {
        if (spaces.Count == 0) throw new ArgumentException("An object-space group cannot be empty.", nameof(spaces));
        OneNoteExtendedGuid objectSpaceId = spaces[0].Id;
        var revisionIds = new HashSet<OneNoteExtendedGuid>();
        var contexts = new HashSet<string>(StringComparer.Ordinal);
        foreach (OneNoteWriteObjectSpace space in spaces) {
            ValidateObjectSpace(space);
            if (!space.Id.Equals(objectSpaceId) || !revisionIds.Add(space.RevisionId) ||
                !contexts.Add(space.ContextId?.ToString() ?? "default")) {
                throw new OneNoteFormatException("ONENOTE_WRITE_OBJECT_SPACE_CONTEXT", "An object-space group contains a mismatched identity or duplicate revision context.");
            }
        }
        OneNoteDesktopFileNodeList manifest = NewList();
        OneNoteDesktopFileNodeList revision = NewList();
        _root.Nodes.Add(OneNoteDesktopFileNode.List(
            OneNoteFileNodeId.ObjectSpaceManifestListReference,
            manifest,
            OneNoteDesktopBinary.Data(stream => OneNoteDesktopBinary.WriteExtendedGuid(stream, objectSpaceId))));
        manifest.Nodes.Add(OneNoteDesktopFileNode.Inline(
            OneNoteFileNodeId.ObjectSpaceManifestListStart,
            OneNoteDesktopBinary.Data(stream => OneNoteDesktopBinary.WriteExtendedGuid(stream, objectSpaceId))));
        manifest.Nodes.Add(OneNoteDesktopFileNode.List(OneNoteFileNodeId.RevisionManifestListReference, revision));

        revision.Nodes.Add(OneNoteDesktopFileNode.Inline(
            OneNoteFileNodeId.RevisionManifestListStart,
            OneNoteDesktopBinary.Data(stream => {
                OneNoteDesktopBinary.WriteExtendedGuid(stream, objectSpaceId);
                FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            })));
        foreach (OneNoteWriteObjectSpace space in spaces) {
            OneNoteFileNodeId startId = _graph.FileKind == OneNoteFileKind.TableOfContents
                ? OneNoteFileNodeId.RevisionManifestStart4
                : space.ContextId == null
                    ? OneNoteFileNodeId.RevisionManifestStart6
                    : OneNoteFileNodeId.RevisionManifestStart7;
            revision.Nodes.Add(OneNoteDesktopFileNode.Inline(startId, CreateRevisionStart(space, _graph.FileKind == OneNoteFileKind.TableOfContents)));

            if (_graph.FileKind == OneNoteFileKind.Section) AddSectionRevision(space, revision);
            else AddTableOfContentsRevision(space, revision);
            revision.Nodes.Add(OneNoteDesktopFileNode.Inline(OneNoteFileNodeId.RevisionManifestEnd));
        }
    }

    private void AddSectionRevision(OneNoteWriteObjectSpace space, OneNoteDesktopFileNodeList revision) {
        OneNoteDesktopFileNodeList group = NewList();
        OneNoteExtendedGuid groupId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 20);
        IReadOnlyDictionary<OneNoteExtendedGuid, uint> referenceCounts = CountReferences(space);
        revision.Nodes.Add(OneNoteDesktopFileNode.List(
            OneNoteFileNodeId.ObjectGroupListReference,
            group,
            OneNoteDesktopBinary.Data(stream => OneNoteDesktopBinary.WriteExtendedGuid(stream, groupId))));
        revision.Nodes.Add(OneNoteDesktopFileNode.Data(
            OneNoteFileNodeId.ObjectInfoDependencyOverrides,
            NilTarget.Instance,
            CreateDependencyOverrides(referenceCounts, space.Objects, _graph.FileKind)));

        IReadOnlyDictionary<Guid, uint> globalIds = BuildGlobalIds(space.Objects);
        AddObjectGroup(group, groupId, space.Objects, globalIds, referenceCounts);
        foreach (KeyValuePair<uint, OneNoteExtendedGuid> root in space.Roots.OrderBy(item => item.Key)) {
            revision.Nodes.Add(OneNoteDesktopFileNode.Inline(
                OneNoteFileNodeId.RootObjectReference3,
                OneNoteDesktopBinary.Data(stream => {
                    OneNoteDesktopBinary.WriteExtendedGuid(stream, root.Value);
                    FssHttpStreamObjectWriter.WriteUInt32(stream, root.Key);
                })));
        }
    }

    private void AddTableOfContentsRevision(OneNoteWriteObjectSpace space, OneNoteDesktopFileNodeList revision) {
        foreach (OneNoteWriteObject item in space.Objects) {
            if (item.Blob != null || item.Jcid != OneNoteSchema.JcidPropertyContainer) {
                throw new OneNoteFormatException("ONENOTE_WRITE_TOC_OBJECT", "A .onetoc2 revision can only contain property-container objects without file data.");
            }
        }
        IReadOnlyDictionary<Guid, uint> globalIds = BuildGlobalIds(space.Objects);
        IReadOnlyDictionary<OneNoteExtendedGuid, uint> referenceCounts = CountReferences(space);
        revision.Nodes.Add(OneNoteDesktopFileNode.Inline(OneNoteFileNodeId.GlobalIdTableStart, new byte[] { 0 }));
        AddGlobalIdEntries(revision, globalIds);
        revision.Nodes.Add(OneNoteDesktopFileNode.Inline(OneNoteFileNodeId.GlobalIdTableEnd));
        revision.Nodes.Add(OneNoteDesktopFileNode.Inline(
            OneNoteFileNodeId.DataSignatureGroupDefinition,
            OneNoteDesktopBinary.Data(stream => OneNoteDesktopBinary.WriteExtendedGuid(stream, NewExtendedGuid()))));
        foreach (OneNoteWriteObject item in space.Objects) {
            AddTableOfContentsObject(revision, item, globalIds, referenceCounts[item.Id]);
        }
        revision.Nodes.Add(OneNoteDesktopFileNode.Data(
            OneNoteFileNodeId.ObjectInfoDependencyOverrides,
            NilTarget.Instance,
            CreateDependencyOverrides(referenceCounts, space.Objects, _graph.FileKind)));
        foreach (KeyValuePair<uint, OneNoteExtendedGuid> root in space.Roots.OrderBy(item => item.Key)) {
            revision.Nodes.Add(OneNoteDesktopFileNode.Inline(
                OneNoteFileNodeId.RootObjectReference2,
                OneNoteDesktopBinary.Data(stream => {
                    FssHttpStreamObjectWriter.WriteUInt32(stream, OneNoteDesktopBinary.CompactId(root.Value, globalIds));
                    FssHttpStreamObjectWriter.WriteUInt32(stream, root.Key);
                })));
        }
    }

    private void AddObjectGroup(
        OneNoteDesktopFileNodeList group,
        OneNoteExtendedGuid groupId,
        IEnumerable<OneNoteWriteObject> objects,
        IReadOnlyDictionary<Guid, uint> globalIds,
        IReadOnlyDictionary<OneNoteExtendedGuid, uint> referenceCounts) {
        group.Nodes.Add(OneNoteDesktopFileNode.Inline(
            OneNoteFileNodeId.ObjectGroupStart,
            OneNoteDesktopBinary.Data(stream => OneNoteDesktopBinary.WriteExtendedGuid(stream, groupId))));
        group.Nodes.Add(OneNoteDesktopFileNode.Inline(OneNoteFileNodeId.GlobalIdTableStart2));
        AddGlobalIdEntries(group, globalIds);
        group.Nodes.Add(OneNoteDesktopFileNode.Inline(OneNoteFileNodeId.GlobalIdTableEnd));
        group.Nodes.Add(OneNoteDesktopFileNode.Inline(
            OneNoteFileNodeId.DataSignatureGroupDefinition,
            OneNoteDesktopBinary.Data(stream => OneNoteDesktopBinary.WriteExtendedGuid(stream, NewExtendedGuid()))));
        foreach (OneNoteWriteObject item in objects) {
            uint referenceCount = referenceCounts[item.Id];
            if (item.Blob != null) AddFileDataObject(group, item, globalIds, referenceCount);
            else AddSectionObject(group, item, globalIds, referenceCount);
        }
        group.Nodes.Add(OneNoteDesktopFileNode.Inline(OneNoteFileNodeId.ObjectGroupEnd));
    }

    private void AddSectionObject(
        OneNoteDesktopFileNodeList group,
        OneNoteWriteObject item,
        IReadOnlyDictionary<Guid, uint> globalIds,
        uint referenceCount) {
        byte[] propertyData = OneNotePropertySetWriter.WriteDesktop(item.Properties, globalIds);
        var chunk = AddData(propertyData);
        OneNoteWriteProperty[] flattenedProperties = OneNotePropertySetWriter.EnumerateProperties(item.Properties).ToArray();
        bool hasObjectReferences = flattenedProperties.Any(property => property.ReferenceKind == OneNoteWriteReferenceKind.Object && property.References.Count > 0);
        bool hasObjectSpaceReferences = flattenedProperties.Any(property => property.ReferenceKind == OneNoteWriteReferenceKind.ObjectSpace && property.References.Count > 0);
        bool isReadOnly = (item.Jcid & 0x00100000U) != 0;
        byte[]? readOnlyHash = null;
        if (isReadOnly) {
            // MS-ONESTORE 2.5.29/2.5.30 mandates MD5 over the referenced, unencrypted object data.
            using (MD5 md5 = MD5.Create()) readOnlyHash = md5.ComputeHash(propertyData);
        }
        byte[] body = OneNoteDesktopBinary.Data(stream => {
            FssHttpStreamObjectWriter.WriteUInt32(stream, OneNoteDesktopBinary.CompactId(item.Id, globalIds));
            FssHttpStreamObjectWriter.WriteUInt32(stream, item.Jcid);
            stream.WriteByte((byte)((hasObjectReferences ? 1 : 0) | (hasObjectSpaceReferences ? 2 : 0)));
            if (referenceCount <= byte.MaxValue) stream.WriteByte((byte)referenceCount);
            else FssHttpStreamObjectWriter.WriteUInt32(stream, referenceCount);
            if (readOnlyHash != null) stream.Write(readOnlyHash, 0, readOnlyHash.Length);
        });
        OneNoteFileNodeId declarationId;
        if (isReadOnly) {
            declarationId = referenceCount <= byte.MaxValue
                ? OneNoteFileNodeId.ReadOnlyObjectDeclaration2RefCount
                : OneNoteFileNodeId.ReadOnlyObjectDeclaration2LargeRefCount;
        } else {
            declarationId = referenceCount <= byte.MaxValue
                ? OneNoteFileNodeId.ObjectDeclaration2RefCount
                : OneNoteFileNodeId.ObjectDeclaration2LargeRefCount;
        }
        group.Nodes.Add(OneNoteDesktopFileNode.Data(declarationId, chunk, body));
    }

    private void AddTableOfContentsObject(
        OneNoteDesktopFileNodeList revision,
        OneNoteWriteObject item,
        IReadOnlyDictionary<Guid, uint> globalIds,
        uint referenceCount) {
        byte[] propertyData = OneNotePropertySetWriter.WriteDesktop(item.Properties, globalIds);
        var chunk = AddData(propertyData);
        bool hasObjectReferences = OneNotePropertySetWriter.EnumerateProperties(item.Properties)
            .Any(property => property.ReferenceKind == OneNoteWriteReferenceKind.Object && property.References.Count > 0);
        byte[] body = OneNoteDesktopBinary.Data(stream => {
            FssHttpStreamObjectWriter.WriteUInt32(stream, OneNoteDesktopBinary.CompactId(item.Id, globalIds));
            uint declarationBits = 1U | (hasObjectReferences ? 0x00010000U : 0U);
            FssHttpStreamObjectWriter.WriteUInt32(stream, declarationBits);
            FssHttpStreamObjectWriter.WriteUInt16(stream, 0);
            if (referenceCount <= byte.MaxValue) stream.WriteByte((byte)referenceCount);
            else FssHttpStreamObjectWriter.WriteUInt32(stream, referenceCount);
        });
        revision.Nodes.Add(OneNoteDesktopFileNode.Data(
            referenceCount <= byte.MaxValue ? OneNoteFileNodeId.ObjectDeclarationWithRefCount : OneNoteFileNodeId.ObjectDeclarationWithRefCount2,
            chunk,
            body));
    }

    private void AddFileDataObject(
        OneNoteDesktopFileNodeList group,
        OneNoteWriteObject item,
        IReadOnlyDictionary<Guid, uint> globalIds,
        uint referenceCount) {
        if (!item.FileDataId.HasValue || item.FileDataId.Value == Guid.Empty) {
            throw new OneNoteFormatException("ONENOTE_WRITE_FILE_DATA_ID", "A desktop OneNote file-data object has no storage identity.");
        }
        string extension = item.FileExtension ?? string.Empty;
        string reference = "<ifndf>{" + item.FileDataId.Value.ToString("D").ToUpperInvariant() + "}";
        byte[] body = OneNoteDesktopBinary.Data(stream => {
            FssHttpStreamObjectWriter.WriteUInt32(stream, OneNoteDesktopBinary.CompactId(item.Id, globalIds));
            FssHttpStreamObjectWriter.WriteUInt32(stream, item.Jcid);
            if (referenceCount <= byte.MaxValue) stream.WriteByte((byte)referenceCount);
            else FssHttpStreamObjectWriter.WriteUInt32(stream, referenceCount);
            WriteStorageString(stream, reference);
            WriteStorageString(stream, extension);
        });
        group.Nodes.Add(OneNoteDesktopFileNode.Inline(
            referenceCount <= byte.MaxValue ? OneNoteFileNodeId.ObjectDeclarationFileData3RefCount : OneNoteFileNodeId.ObjectDeclarationFileData3LargeRefCount,
            body));

        OneNoteDesktopDataChunk frame = AddData(CreateFileDataFrame(item.Blob!));
        if (_fileDataList == null) _fileDataList = NewList();
        _fileDataList.Nodes.Add(OneNoteDesktopFileNode.Data(
            OneNoteFileNodeId.FileDataStoreObjectReference,
            frame,
            item.FileDataId.Value.ToByteArray()));
    }

    private static byte[] CreateRevisionStart(OneNoteWriteObjectSpace space, bool tableOfContents) => OneNoteDesktopBinary.Data(stream => {
        OneNoteDesktopBinary.WriteExtendedGuid(stream, space.RevisionId);
        OneNoteDesktopBinary.WriteExtendedGuid(stream, space.DependencyId ?? new OneNoteExtendedGuid(Guid.Empty, 0, 20));
        if (tableOfContents) FssHttpStreamObjectWriter.WriteUInt64(stream, 0);
        FssHttpStreamObjectWriter.WriteUInt32(stream, 1);
        FssHttpStreamObjectWriter.WriteUInt16(stream, 0);
        if (space.ContextId != null) OneNoteDesktopBinary.WriteExtendedGuid(stream, space.ContextId);
    });

    private byte[] CreateFileDataFrame(byte[] payload) {
        long footerOffset = OneNoteDesktopBinary.Align8(36L + payload.LongLength);
        long length = checked(footerOffset + 16L);
        if (length > uint.MaxValue || length > int.MaxValue) throw new OneNoteFormatException("ONENOTE_WRITE_FILE_DATA_SIZE", "An embedded OneNote file-data frame is too large.");
        var result = new byte[(int)length];
        using (var stream = new MemoryStream(result, true)) {
            OneNoteDesktopBinary.WriteGuid(stream, FileDataHeader);
            FssHttpStreamObjectWriter.WriteUInt64(stream, (ulong)payload.LongLength);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt64(stream, 0);
            stream.Write(payload, 0, payload.Length);
            stream.Position = footerOffset;
            OneNoteDesktopBinary.WriteGuid(stream, FileDataFooter);
        }
        return result;
    }

    private static byte[] CreateDependencyOverrides(
        IReadOnlyDictionary<OneNoteExtendedGuid, uint> referenceCounts,
        IEnumerable<OneNoteWriteObject> objects,
        OneNoteFileKind fileKind) => OneNoteDesktopBinary.Data(stream => {
            byte[] counts = OneNoteDesktopBinary.Data(countStream => {
                foreach (OneNoteWriteObject item in objects) {
                    FssHttpStreamObjectWriter.WriteUInt32(countStream, referenceCounts[item.Id]);
                }
            });
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt32(stream, OneNoteCrc32.Continue(0, counts, 0, counts.Length, fileKind));
        });

    private static IReadOnlyDictionary<OneNoteExtendedGuid, uint> CountReferences(OneNoteWriteObjectSpace space) {
        var counts = space.Objects.ToDictionary(item => item.Id, _ => 0U);
        foreach (OneNoteWriteObject item in space.Objects) {
            foreach (OneNoteExtendedGuid reference in OneNotePropertySetWriter.EnumerateProperties(item.Properties)
                .Where(property => property.ReferenceKind == OneNoteWriteReferenceKind.Object)
                .SelectMany(property => property.References)) {
                if (counts.TryGetValue(reference, out uint count)) counts[reference] = checked(count + 1U);
            }
        }
        foreach (OneNoteExtendedGuid root in space.Roots.Values) {
            if (!counts.TryGetValue(root, out uint count)) throw new OneNoteFormatException("ONENOTE_WRITE_ROOT_OBJECT", "A revision root does not identify an object in its object space.");
            counts[root] = checked(count + 1U);
        }
        return counts;
    }

    private static IReadOnlyDictionary<Guid, uint> BuildGlobalIds(IEnumerable<OneNoteWriteObject> objects) {
        var identifiers = new List<Guid>();
        foreach (OneNoteWriteObject item in objects) {
            AddIdentifier(identifiers, item.Id.Identifier);
            foreach (OneNoteExtendedGuid reference in OneNotePropertySetWriter.EnumerateProperties(item.Properties).SelectMany(property => property.References)) {
                AddIdentifier(identifiers, reference.Identifier);
            }
        }
        if (identifiers.Count >= 0xFFFFFF) throw new OneNoteFormatException("ONENOTE_WRITE_GLOBAL_ID_LIMIT", "A desktop OneNote global-identification table exceeds the CompactID index range.");
        return identifiers.Select((identifier, index) => new { identifier, index }).ToDictionary(item => item.identifier, item => (uint)item.index);
    }

    private static void AddGlobalIdEntries(OneNoteDesktopFileNodeList list, IReadOnlyDictionary<Guid, uint> globalIds) {
        foreach (KeyValuePair<Guid, uint> entry in globalIds.OrderBy(item => item.Value)) {
            list.Nodes.Add(OneNoteDesktopFileNode.Inline(
                OneNoteFileNodeId.GlobalIdTableEntry,
                OneNoteDesktopBinary.Data(stream => {
                    FssHttpStreamObjectWriter.WriteUInt32(stream, entry.Value);
                    OneNoteDesktopBinary.WriteGuid(stream, entry.Key);
                })));
        }
    }

    private static void AddIdentifier(ICollection<Guid> identifiers, Guid value) {
        if (value == Guid.Empty) throw new OneNoteFormatException("ONENOTE_WRITE_GLOBAL_ID_GUID", "A desktop OneNote CompactID cannot use an empty GUID.");
        if (!identifiers.Contains(value)) identifiers.Add(value);
    }

    private static void WriteStorageString(Stream stream, string value) {
        FssHttpStreamObjectWriter.WriteUInt32(stream, (uint)value.Length);
        byte[] data = Encoding.Unicode.GetBytes(value);
        stream.Write(data, 0, data.Length);
    }

    private static OneNoteExtendedGuid NewExtendedGuid() => new OneNoteExtendedGuid(Guid.NewGuid(), 1, 20);

    private OneNoteDesktopDataChunk AddData(byte[] data) {
        var chunk = new OneNoteDesktopDataChunk(data);
        _data.Add(chunk);
        return chunk;
    }

    private OneNoteDesktopFileNodeList NewList() {
        var list = new OneNoteDesktopFileNodeList(_nextListId++);
        _lists.Add(list);
        return list;
    }

    private static void ValidateObjectSpace(OneNoteWriteObjectSpace space) {
        if (space.Id.Identifier == Guid.Empty || space.Id.Value == 0 || space.RevisionId.Identifier == Guid.Empty || space.RevisionId.Value == 0) {
            throw new OneNoteFormatException("ONENOTE_WRITE_OBJECT_SPACE_ID", "A desktop OneNote object space or revision has an empty identity.");
        }
        var seen = new HashSet<OneNoteExtendedGuid>();
        foreach (OneNoteWriteObject item in space.Objects) {
            if (item.Id.Identifier == Guid.Empty || item.Id.Value == 0 || item.Id.Value > byte.MaxValue || !seen.Add(item.Id)) {
                throw new OneNoteFormatException("ONENOTE_WRITE_OBJECT_ID", "A desktop OneNote object identity is empty, duplicated, or cannot be represented as a CompactID.");
            }
        }
    }

    private sealed class NilTarget : IOneNoteDesktopReferenceTarget {
        internal static readonly NilTarget Instance = new NilTarget();
        private NilTarget() { }
        public ulong Offset { get => ulong.MaxValue; set => throw new NotSupportedException(); }
        public uint Length => 0;
    }
}
