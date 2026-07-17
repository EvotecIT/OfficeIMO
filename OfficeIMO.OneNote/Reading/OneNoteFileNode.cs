namespace OfficeIMO.OneNote;

/// <summary>
/// Known MS-ONESTORE file-node identifiers. Unknown values remain available through
/// <see cref="OneNoteFileNode.RawId"/>.
/// </summary>
public enum OneNoteFileNodeId : ushort {
    /// <summary>An identifier not currently mapped by OfficeIMO.OneNote.</summary>
    Unknown = 0,
    /// <summary>Declares the root object space.</summary>
    ObjectSpaceManifestRoot = 0x004,
    /// <summary>References an object-space manifest list.</summary>
    ObjectSpaceManifestListReference = 0x008,
    /// <summary>Begins an object-space manifest list.</summary>
    ObjectSpaceManifestListStart = 0x00C,
    /// <summary>References a revision-manifest list.</summary>
    RevisionManifestListReference = 0x010,
    /// <summary>Begins a revision-manifest list.</summary>
    RevisionManifestListStart = 0x014,
    /// <summary>Begins a version-4 revision manifest.</summary>
    RevisionManifestStart4 = 0x01B,
    /// <summary>Ends a revision manifest.</summary>
    RevisionManifestEnd = 0x01C,
    /// <summary>Begins a version-6 revision manifest.</summary>
    RevisionManifestStart6 = 0x01E,
    /// <summary>Begins a version-7 revision manifest.</summary>
    RevisionManifestStart7 = 0x01F,
    /// <summary>Begins a compact global-identification table.</summary>
    GlobalIdTableStart = 0x021,
    /// <summary>Begins a global-identification table.</summary>
    GlobalIdTableStart2 = 0x022,
    /// <summary>Declares a global-identification table entry.</summary>
    GlobalIdTableEntry = 0x024,
    /// <summary>Maps one dependency global-identification entry.</summary>
    GlobalIdTableEntry2 = 0x025,
    /// <summary>Maps a range of dependency global-identification entries.</summary>
    GlobalIdTableEntry3 = 0x026,
    /// <summary>Ends a global-identification table.</summary>
    GlobalIdTableEnd = 0x028,
    /// <summary>Declares an object with a compact reference count.</summary>
    ObjectDeclarationWithRefCount = 0x02D,
    /// <summary>Declares an object with a large reference count.</summary>
    ObjectDeclarationWithRefCount2 = 0x02E,
    /// <summary>Revises an object with a compact reference count.</summary>
    ObjectRevisionWithRefCount = 0x041,
    /// <summary>Revises an object with a large reference count.</summary>
    ObjectRevisionWithRefCount2 = 0x042,
    /// <summary>References a root object using a compact identifier.</summary>
    RootObjectReference2 = 0x059,
    /// <summary>References a root object using an extended identifier.</summary>
    RootObjectReference3 = 0x05A,
    /// <summary>Declares a revision role.</summary>
    RevisionRoleDeclaration = 0x05C,
    /// <summary>Declares a revision role and context.</summary>
    RevisionRoleAndContextDeclaration = 0x05D,
    /// <summary>Declares a file-data object.</summary>
    ObjectDeclarationFileData3RefCount = 0x072,
    /// <summary>Declares a file-data object with a large reference count.</summary>
    ObjectDeclarationFileData3LargeRefCount = 0x073,
    /// <summary>Declares an object-data encryption key.</summary>
    ObjectDataEncryptionKeyV2 = 0x07C,
    /// <summary>Overrides object dependency information.</summary>
    ObjectInfoDependencyOverrides = 0x084,
    /// <summary>Defines a data-signature group.</summary>
    DataSignatureGroupDefinition = 0x08C,
    /// <summary>References the file-data store list.</summary>
    FileDataStoreListReference = 0x090,
    /// <summary>References one file-data store object.</summary>
    FileDataStoreObjectReference = 0x094,
    /// <summary>Declares a version-2 object with a compact reference count.</summary>
    ObjectDeclaration2RefCount = 0x0A4,
    /// <summary>Declares a version-2 object with a large reference count.</summary>
    ObjectDeclaration2LargeRefCount = 0x0A5,
    /// <summary>References an object-group list.</summary>
    ObjectGroupListReference = 0x0B0,
    /// <summary>Begins an object group.</summary>
    ObjectGroupStart = 0x0B4,
    /// <summary>Ends an object group.</summary>
    ObjectGroupEnd = 0x0B8,
    /// <summary>Declares a read-only object.</summary>
    ReadOnlyObjectDeclaration2RefCount = 0x0C4,
    /// <summary>Declares a read-only object with a large reference count.</summary>
    ReadOnlyObjectDeclaration2LargeRefCount = 0x0C5,
    /// <summary>Terminates a non-final file-node-list fragment.</summary>
    ChunkTerminator = 0x0FF
}

/// <summary>How a file node uses its leading chunk reference.</summary>
public enum OneNoteFileNodeBaseType : byte {
    /// <summary>The node has no leading chunk reference.</summary>
    Inline = 0,
    /// <summary>The node references data.</summary>
    DataReference = 1,
    /// <summary>The node references another file-node list.</summary>
    FileNodeListReference = 2
}

/// <summary>A decoded variable-width chunk reference stored in a file node.</summary>
public sealed class OneNoteFileNodeChunkReference {
    internal OneNoteFileNodeChunkReference(ulong offset, ulong length, bool isNil, int encodedLength) {
        Offset = offset;
        Length = length;
        IsNil = isNil;
        EncodedLength = encodedLength;
    }

    /// <summary>Referenced offset from the start of the revision-store file.</summary>
    public ulong Offset { get; }

    /// <summary>Referenced byte count.</summary>
    public ulong Length { get; }

    /// <summary>Whether the encoded reference is the format's nil sentinel.</summary>
    public bool IsNil { get; }

    /// <summary>Number of bytes occupied by the variable-width encoded reference.</summary>
    public int EncodedLength { get; }

    /// <summary>Whether both decoded fields are zero.</summary>
    public bool IsZero => Offset == 0 && Length == 0;

    /// <inheritdoc />
    public override string ToString() => IsNil ? "nil" : Offset.ToString(System.Globalization.CultureInfo.InvariantCulture) + "+" + Length.ToString(System.Globalization.CultureInfo.InvariantCulture);
}

/// <summary>A decoded MS-ONESTORE file node.</summary>
public sealed class OneNoteFileNode {
    internal OneNoteFileNode(
        ushort rawId,
        int size,
        byte stpFormat,
        byte cbFormat,
        OneNoteFileNodeBaseType baseType,
        long fileOffset,
        OneNoteFileNodeChunkReference? chunkReference,
        byte[] encodedData) {
        RawId = rawId;
        Size = size;
        StpFormat = stpFormat;
        CbFormat = cbFormat;
        BaseType = baseType;
        FileOffset = fileOffset;
        ChunkReference = chunkReference;
        EncodedData = OneNoteBinaryPayload.FromBytes(encodedData);
    }

    /// <summary>Raw 10-bit file-node identifier.</summary>
    public ushort RawId { get; }

    /// <summary>Known identifier mapping, or <see cref="OneNoteFileNodeId.Unknown"/>.</summary>
    public OneNoteFileNodeId Id => Enum.IsDefined(typeof(OneNoteFileNodeId), RawId) ? (OneNoteFileNodeId)RawId : OneNoteFileNodeId.Unknown;

    /// <summary>Total encoded node size, including its four-byte header.</summary>
    public int Size { get; }

    /// <summary>Encoded file-pointer format selector.</summary>
    public byte StpFormat { get; }

    /// <summary>Encoded byte-count format selector.</summary>
    public byte CbFormat { get; }

    /// <summary>Whether the node stores inline data or a data/list reference.</summary>
    public OneNoteFileNodeBaseType BaseType { get; }

    /// <summary>Absolute byte offset of the node header.</summary>
    public long FileOffset { get; }

    /// <summary>Leading decoded chunk reference for base types 1 and 2.</summary>
    public OneNoteFileNodeChunkReference? ChunkReference { get; }

    /// <summary>Complete encoded <c>fnd</c> payload, including any leading chunk reference.</summary>
    public OneNoteBinaryPayload EncodedData { get; }

    /// <summary>Decoded child list when this node has file-node-list base type.</summary>
    public OneNoteFileNodeList? ReferencedFileNodeList { get; internal set; }
}
