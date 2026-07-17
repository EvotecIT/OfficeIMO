namespace OfficeIMO.OneNote;

internal enum OneNoteWriteReferenceKind { Object, ObjectSpace, Context }

internal sealed class OneNoteWriteProperty {
    internal OneNoteWriteProperty(
        uint rawId,
        byte[]? data = null,
        ulong? scalar = null,
        bool? boolean = null,
        IEnumerable<OneNoteExtendedGuid>? references = null,
        OneNoteWriteReferenceKind referenceKind = OneNoteWriteReferenceKind.Object,
        IEnumerable<IReadOnlyList<OneNoteWriteProperty>>? childPropertySets = null,
        uint? childPropertyId = null,
        bool preserveRawId = false) {
        RawId = preserveRawId
            ? rawId
            : boolean.HasValue && boolean.Value ? rawId | 0x80000000U : rawId & 0x7FFFFFFFU;
        Data = data;
        Scalar = scalar;
        References = references?.ToArray() ?? Array.Empty<OneNoteExtendedGuid>();
        ReferenceKind = referenceKind;
        ChildPropertySets = childPropertySets?.ToArray() ?? Array.Empty<IReadOnlyList<OneNoteWriteProperty>>();
        ChildPropertyId = childPropertyId;
    }
    internal uint RawId { get; }
    internal byte[]? Data { get; }
    internal ulong? Scalar { get; }
    internal IReadOnlyList<OneNoteExtendedGuid> References { get; }
    internal OneNoteWriteReferenceKind ReferenceKind { get; }
    internal IReadOnlyList<IReadOnlyList<OneNoteWriteProperty>> ChildPropertySets { get; }
    internal uint? ChildPropertyId { get; }
}

internal sealed class OneNoteWriteObject {
    internal OneNoteWriteObject(
        OneNoteExtendedGuid id,
        uint jcid,
        IEnumerable<OneNoteWriteProperty>? properties = null,
        byte[]? blob = null,
        Guid? fileDataId = null,
        string? fileExtension = null) {
        Id = id;
        Jcid = jcid;
        Properties = properties?.ToArray() ?? Array.Empty<OneNoteWriteProperty>();
        Blob = blob;
        FileDataId = fileDataId;
        FileExtension = fileExtension;
    }
    internal OneNoteExtendedGuid Id { get; }
    internal uint Jcid { get; }
    internal IReadOnlyList<OneNoteWriteProperty> Properties { get; }
    internal byte[]? Blob { get; }
    internal Guid? FileDataId { get; }
    internal string? FileExtension { get; }
}

internal sealed class OneNoteWriteObjectSpace {
    internal OneNoteWriteObjectSpace(
        OneNoteExtendedGuid id,
        OneNoteExtendedGuid revisionId,
        OneNoteExtendedGuid? contextId = null,
        OneNoteExtendedGuid? dependencyId = null) {
        Id = id;
        RevisionId = revisionId;
        ContextId = contextId;
        DependencyId = dependencyId;
    }
    internal OneNoteExtendedGuid Id { get; }
    internal OneNoteExtendedGuid RevisionId { get; }
    internal OneNoteExtendedGuid? ContextId { get; }
    internal OneNoteExtendedGuid? DependencyId { get; }
    internal IList<OneNoteWriteObject> Objects { get; } = new List<OneNoteWriteObject>();
    internal IDictionary<uint, OneNoteExtendedGuid> Roots { get; } = new Dictionary<uint, OneNoteExtendedGuid>();
}

internal sealed class OneNoteWriteGraph {
    internal OneNoteWriteGraph(Guid fileId, OneNoteFileKind fileKind, OneNoteExtendedGuid rootObjectSpaceId, Guid ancestorId, uint fileNameCrc) {
        FileId = fileId;
        FileKind = fileKind;
        RootObjectSpaceId = rootObjectSpaceId;
        AncestorId = ancestorId;
        FileNameCrc = fileNameCrc;
    }
    internal Guid FileId { get; }
    internal OneNoteFileKind FileKind { get; }
    internal OneNoteExtendedGuid RootObjectSpaceId { get; }
    internal Guid AncestorId { get; }
    internal uint FileNameCrc { get; }
    internal IList<OneNoteWriteObjectSpace> ObjectSpaces { get; } = new List<OneNoteWriteObjectSpace>();
}

internal sealed class OneNoteWriteIdFactory {
    internal OneNoteExtendedGuid New() => new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
}

internal sealed class OneNoteTocWriteEntry {
    internal OneNoteTocWriteEntry(Guid id, string name, uint order, uint? colorArgb) {
        Id = id;
        Name = name;
        Order = order;
        ColorArgb = colorArgb;
    }

    internal Guid Id { get; }
    internal string Name { get; }
    internal uint Order { get; }
    internal uint? ColorArgb { get; }
}
