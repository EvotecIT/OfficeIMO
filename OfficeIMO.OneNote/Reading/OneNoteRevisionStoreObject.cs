namespace OfficeIMO.OneNote;

/// <summary>Exact property representation encoded by an MS-ONESTORE PropertyID.</summary>
public enum OneNotePropertyType : byte {
    /// <summary>An unknown or unsupported property representation.</summary>
    Unknown = 0,
    /// <summary>No value bytes.</summary>
    NoData = 0x01,
    /// <summary>A Boolean encoded in the PropertyID.</summary>
    Boolean = 0x02,
    /// <summary>One byte in the property-data stream.</summary>
    Byte = 0x03,
    /// <summary>Two bytes in the property-data stream.</summary>
    UInt16 = 0x04,
    /// <summary>Four bytes in the property-data stream.</summary>
    UInt32 = 0x05,
    /// <summary>Eight bytes in the property-data stream.</summary>
    UInt64 = 0x06,
    /// <summary>A four-byte length followed by data.</summary>
    LengthPrefixedData = 0x07,
    /// <summary>One object identifier from the OID stream.</summary>
    ObjectId = 0x08,
    /// <summary>An array of identifiers from the OID stream.</summary>
    ObjectIdArray = 0x09,
    /// <summary>One object-space identifier from the OSID stream.</summary>
    ObjectSpaceId = 0x0A,
    /// <summary>An array of identifiers from the OSID stream.</summary>
    ObjectSpaceIdArray = 0x0B,
    /// <summary>One context identifier from the context stream.</summary>
    ContextId = 0x0C,
    /// <summary>An array of identifiers from the context stream.</summary>
    ContextIdArray = 0x0D,
    /// <summary>An array of child property sets.</summary>
    PropertySetArray = 0x10,
    /// <summary>One child property set.</summary>
    PropertySet = 0x11
}

/// <summary>A decoded property and its loss-aware encoded value.</summary>
public sealed class OneNotePropertyValue {
    internal OneNotePropertyValue(uint rawPropertyId, int ordinal) {
        RawPropertyId = rawPropertyId;
        Ordinal = ordinal;
    }

    /// <summary>Complete encoded PropertyID, including type and inline Boolean bits.</summary>
    public uint RawPropertyId { get; }

    /// <summary>Twenty-six-bit semantic property identifier.</summary>
    public uint Id => RawPropertyId & 0x03FFFFFFU;

    /// <summary>Property representation.</summary>
    public OneNotePropertyType Type {
        get {
            byte value = (byte)((RawPropertyId >> 26) & 0x1FU);
            return Enum.IsDefined(typeof(OneNotePropertyType), value) ? (OneNotePropertyType)value : OneNotePropertyType.Unknown;
        }
    }

    /// <summary>Original property order.</summary>
    public int Ordinal { get; }

    /// <summary>Inline Boolean value when <see cref="Type"/> is Boolean.</summary>
    public bool? BooleanValue { get; internal set; }

    /// <summary>Unsigned scalar value for fixed-width numeric representations.</summary>
    public ulong? ScalarValue { get; internal set; }

    /// <summary>Length-prefixed or otherwise retained encoded value bytes.</summary>
    public OneNoteBinaryPayload? Data { get; internal set; }

    /// <summary>Resolved object, object-space, or context identifiers.</summary>
    public IReadOnlyList<OneNoteExtendedGuid> ReferencedIds { get; internal set; } = Array.Empty<OneNoteExtendedGuid>();

    /// <summary>Nested property sets.</summary>
    public IReadOnlyList<OneNotePropertySet> ChildPropertySets { get; internal set; } = Array.Empty<OneNotePropertySet>();

    /// <summary>Property identifier used for each element of a property-set array.</summary>
    public uint? ChildPropertyId { get; internal set; }
}

/// <summary>A decoded property set in source order.</summary>
public sealed class OneNotePropertySet {
    internal OneNotePropertySet(IReadOnlyList<OneNotePropertyValue> properties, int encodedLength) {
        Properties = properties;
        EncodedLength = encodedLength;
    }

    /// <summary>Properties in encoded order.</summary>
    public IReadOnlyList<OneNotePropertyValue> Properties { get; }

    /// <summary>Bytes consumed by this property set.</summary>
    public int EncodedLength { get; }

    /// <summary>Finds the last property with the complete PropertyID value.</summary>
    public OneNotePropertyValue? Find(uint rawPropertyId) => Properties.LastOrDefault(property => property.RawPropertyId == rawPropertyId);
}

/// <summary>Decoded JCID object classification.</summary>
public sealed class OneNoteJcid {
    internal OneNoteJcid(uint value) { Value = value; }

    /// <summary>Complete JCID value.</summary>
    public uint Value { get; }

    /// <summary>Object type index.</summary>
    public ushort Index => (ushort)(Value & 0xFFFFU);

    /// <summary>Whether the object carries synchronization encryption data.</summary>
    public bool IsBinary => (Value & 0x00010000U) != 0;

    /// <summary>Whether the object data is a property set.</summary>
    public bool IsPropertySet => (Value & 0x00020000U) != 0;

    /// <summary>Whether the object is a graph node.</summary>
    public bool IsGraphNode => (Value & 0x00040000U) != 0;

    /// <summary>Whether the object is file data.</summary>
    public bool IsFileData => (Value & 0x00080000U) != 0;

    /// <summary>Whether revisions must retain identical data.</summary>
    public bool IsReadOnly => (Value & 0x00100000U) != 0;
}

/// <summary>A declared or revised object from a revision manifest.</summary>
public sealed class OneNoteRevisionStoreObject {
    internal OneNoteRevisionStoreObject(OneNoteExtendedGuid id, OneNoteJcid jcid, OneNoteFileNode declarationNode) {
        Id = id;
        Jcid = jcid;
        DeclarationNode = declarationNode;
    }

    /// <summary>Stable object identity.</summary>
    public OneNoteExtendedGuid Id { get; }

    /// <summary>Object class and representation flags.</summary>
    public OneNoteJcid Jcid { get; internal set; }

    /// <summary>Reference count declared by this object version.</summary>
    public uint ReferenceCount { get; internal set; }

    /// <summary>Revision that contains this declaration.</summary>
    public OneNoteExtendedGuid? RevisionId { get; internal set; }

    /// <summary>Whether the record revises an earlier declaration.</summary>
    public bool IsRevision { get; internal set; }

    /// <summary>Source declaration node.</summary>
    public OneNoteFileNode DeclarationNode { get; }

    /// <summary>Decoded property-set data.</summary>
    public OneNotePropertySet? PropertySet { get; internal set; }

    /// <summary>Exact encoded property-stream bytes retained for loss-aware round trips.</summary>
    public OneNoteBinaryPayload? RawPropertyData { get; internal set; }

    /// <summary>File-data reference such as <c>&lt;ifndf&gt;{guid}</c>.</summary>
    public string? FileDataReference { get; internal set; }

    /// <summary>File extension declared for a file-data object.</summary>
    public string? FileExtension { get; internal set; }
}

/// <summary>A revision manifest and its root-role choices.</summary>
public sealed class OneNoteRevisionManifest {
    private readonly List<OneNoteRevisionRoleAssociation> _roleAssociations = new List<OneNoteRevisionRoleAssociation>();

    internal OneNoteRevisionManifest(OneNoteExtendedGuid id) { Id = id; }

    /// <summary>Revision identity.</summary>
    public OneNoteExtendedGuid Id { get; }

    /// <summary>Object space revised by this manifest.</summary>
    public OneNoteExtendedGuid? ObjectSpaceId { get; internal set; }

    /// <summary>Dependency revision identity.</summary>
    public OneNoteExtendedGuid? DependencyId { get; internal set; }

    /// <summary>Context identity for a version-7 manifest.</summary>
    public OneNoteExtendedGuid? ContextId { get; internal set; }

    /// <summary>Revision role.</summary>
    public uint Role { get; internal set; }

    /// <summary>
    /// Context and revision-role labels associated with this revision in source order. Later
    /// associations for the same label supersede earlier ones.
    /// </summary>
    public IReadOnlyList<OneNoteRevisionRoleAssociation> RoleAssociations => _roleAssociations;

    /// <summary>Whether the revision's object data is encrypted.</summary>
    public bool IsEncrypted { get; internal set; }

    /// <summary>Root objects selected by this revision.</summary>
    public IList<OneNoteRootObjectReference> RootObjects { get; } = new List<OneNoteRootObjectReference>();

    internal void AddRoleAssociation(OneNoteExtendedGuid? contextId, uint role, int sourceOrder) {
        _roleAssociations.Add(new OneNoteRevisionRoleAssociation(contextId, role, sourceOrder));
    }
}

/// <summary>A context and revision-role label associated with a revision manifest.</summary>
public sealed class OneNoteRevisionRoleAssociation {
    internal OneNoteRevisionRoleAssociation(OneNoteExtendedGuid? contextId, uint role, int sourceOrder) {
        ContextId = contextId;
        Role = role;
        SourceOrder = sourceOrder;
    }

    /// <summary>Context identity, or <see langword="null"/> for the default context.</summary>
    public OneNoteExtendedGuid? ContextId { get; }

    /// <summary>Revision-role label.</summary>
    public uint Role { get; }

    internal int SourceOrder { get; }
}

/// <summary>A root object selected for one revision role.</summary>
public sealed class OneNoteRootObjectReference {
    internal OneNoteRootObjectReference(OneNoteExtendedGuid objectId, uint role) {
        ObjectId = objectId;
        Role = role;
    }

    /// <summary>Root object identity.</summary>
    public OneNoteExtendedGuid ObjectId { get; }

    /// <summary>MS-ONE root role.</summary>
    public uint Role { get; }
}

/// <summary>An embedded FileDataStoreObject payload.</summary>
public sealed class OneNoteFileDataStoreObject {
    internal OneNoteFileDataStoreObject(Guid id, OneNoteBinaryPayload payload) {
        Id = id;
        Payload = payload;
    }

    /// <summary>File-data store identity.</summary>
    public Guid Id { get; }

    /// <summary>Embedded payload without FileDataStoreObject framing.</summary>
    public OneNoteBinaryPayload Payload { get; }
}
