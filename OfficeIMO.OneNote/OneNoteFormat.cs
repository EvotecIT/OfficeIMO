namespace OfficeIMO.OneNote;

/// <summary>
/// Identifies the logical OneNote artifact represented by a file.
/// </summary>
public enum OneNoteFileKind {
    /// <summary>The artifact could not be classified.</summary>
    Unknown = 0,
    /// <summary>A OneNote section stored in a <c>.one</c> file.</summary>
    Section = 1,
    /// <summary>A OneNote notebook table of contents stored in a <c>.onetoc2</c> file.</summary>
    TableOfContents = 2,
    /// <summary>A packaged OneNote notebook stored in a <c>.onepkg</c> file.</summary>
    NotebookPackage = 3
}

/// <summary>
/// Identifies the physical encoding used by a OneNote artifact.
/// </summary>
public enum OneNoteStorageFormat {
    /// <summary>The physical encoding could not be classified.</summary>
    Unknown = 0,
    /// <summary>The desktop MS-ONESTORE revision-store encoding.</summary>
    RevisionStore = 1,
    /// <summary>The MS-FSSHTTPB data-element package encoding used for server-backed files.</summary>
    FileSynchronizationPackage = 2,
    /// <summary>The notebook archive encoding used by <c>.onepkg</c> files.</summary>
    NotebookPackage = 3
}

/// <summary>
/// Identifies a property representation used by MS-ONESTORE property sets.
/// </summary>
public enum OneNotePropertyValueType {
    /// <summary>The property representation is unknown.</summary>
    Unknown = 0,
    /// <summary>A single byte value.</summary>
    Byte = 1,
    /// <summary>A two-byte value.</summary>
    UInt16 = 2,
    /// <summary>A four-byte value.</summary>
    UInt32 = 3,
    /// <summary>An eight-byte value.</summary>
    UInt64 = 4,
    /// <summary>A length-prefixed byte sequence.</summary>
    Blob = 5,
    /// <summary>An object identifier.</summary>
    ObjectId = 6,
    /// <summary>An object-space identifier.</summary>
    ObjectSpaceId = 7,
    /// <summary>A context identifier.</summary>
    ContextId = 8,
    /// <summary>An array of object identifiers.</summary>
    ObjectIdArray = 9,
    /// <summary>An array of object-space identifiers.</summary>
    ObjectSpaceIdArray = 10,
    /// <summary>An array of context identifiers.</summary>
    ContextIdArray = 11,
    /// <summary>A property with no separate value bytes.</summary>
    NoData = 12,
    /// <summary>A Boolean encoded directly in the property identifier.</summary>
    Boolean = 13,
    /// <summary>A nested property set.</summary>
    PropertySet = 14,
    /// <summary>An array of nested property sets.</summary>
    PropertySetArray = 15
}
