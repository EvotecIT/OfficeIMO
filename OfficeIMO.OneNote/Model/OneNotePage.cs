namespace OfficeIMO.OneNote;

/// <summary>
/// A OneNote page, including subpage level, content, revisions, and conflict metadata.
/// </summary>
public sealed class OneNotePage {
    internal OneNotePagePreservationIds PreservationIds { get; } = new OneNotePagePreservationIds();

    /// <summary>Page object-space identity when available. Serialization assigns and retains an identity for a new page.</summary>
    public OneNoteExtendedGuid? Id { get; set; }

    /// <summary>Revision context identity for a historical page snapshot. Serialization assigns it when needed.</summary>
    public OneNoteExtendedGuid? RevisionContextId { get; set; }

    /// <summary>Page title.</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>Zero-based subpage nesting level.</summary>
    public int Level { get; set; }

    /// <summary>Page creation time.</summary>
    public DateTime? CreatedUtc { get; set; }

    /// <summary>Most recent page modification time.</summary>
    public DateTime? LastModifiedUtc { get; set; }

    /// <summary>Original author.</summary>
    public string? OriginalAuthor { get; set; }

    /// <summary>Most recent author.</summary>
    public string? MostRecentAuthor { get; set; }

    /// <summary>Whether this is a conflict page.</summary>
    public bool IsConflictPage { get; set; }

    /// <summary>Whether this is a version-history page.</summary>
    public bool IsVersionHistoryPage { get; set; }

    /// <summary>Whether the page is marked as deleted content.</summary>
    public bool IsDeleted { get; set; }

    /// <summary>Optional width in OneNote layout units.</summary>
    public double? Width { get; set; }

    /// <summary>Optional height in OneNote layout units.</summary>
    public double? Height { get; set; }

    /// <summary>Page-level outlines in source order.</summary>
    public IList<OneNoteOutline> Outlines { get; } = new List<OneNoteOutline>();

    /// <summary>Content placed directly on the page outside outlines.</summary>
    public IList<OneNoteElement> DirectContent { get; } = new List<OneNoteElement>();

    /// <summary>Conflict pages associated with this page.</summary>
    public IList<OneNotePage> ConflictPages { get; } = new List<OneNotePage>();

    /// <summary>Version-history snapshots associated with this page.</summary>
    public IList<OneNotePage> VersionHistory { get; } = new List<OneNotePage>();

    /// <summary>Revision metadata associated with the page object space.</summary>
    public IList<OneNoteRevision> Revisions { get; } = new List<OneNoteRevision>();

    /// <summary>Opaque page objects preserved for loss-aware writing.</summary>
    public IList<OneNoteOpaqueObject> UnknownObjects { get; } = new List<OneNoteOpaqueObject>();

    /// <summary>Diagnostics produced while loading the page.</summary>
    public IList<OneNoteDiagnostic> Diagnostics { get; } = new List<OneNoteDiagnostic>();
}

internal sealed class OneNotePagePreservationIds {
    internal OneNoteExtendedGuid? ManifestId { get; set; }
    internal OneNoteExtendedGuid? MetadataId { get; set; }
    internal OneNoteExtendedGuid? RevisionMetadataId { get; set; }
    internal OneNoteExtendedGuid? PageNodeId { get; set; }
    internal OneNoteExtendedGuid? TitleNodeId { get; set; }
    internal OneNoteExtendedGuid? TitleOutlineId { get; set; }
    internal OneNoteExtendedGuid? TitleElementId { get; set; }
    internal OneNoteExtendedGuid? TitleTextId { get; set; }
    internal OneNoteExtendedGuid? VersionProxyId { get; set; }
}

/// <summary>
/// Author metadata attached to a page or content object.
/// </summary>
public sealed class OneNoteAuthor {
    /// <summary>Display name.</summary>
    public string? Name { get; set; }

    /// <summary>Native author-object identity. Serialization assigns and retains it for new author metadata.</summary>
    public OneNoteExtendedGuid? ObjectId { get; set; }
}

/// <summary>
/// A note tag or task marker associated with content.
/// </summary>
public sealed class OneNoteTag {
    /// <summary>Tag definition identity. Serialization assigns and retains it for a new non-task tag.</summary>
    public OneNoteExtendedGuid? DefinitionId { get; set; }

    /// <summary>MS-ONE action-item identity (0-99 for normal tags and 100-105 for task tags).</summary>
    public uint? ActionItemType { get; set; }

    /// <summary>Tag label.</summary>
    public string? Label { get; set; }

    /// <summary>MS-ONE note-tag shape value.</summary>
    public uint? Shape { get; set; }

    /// <summary>Whether the tag is a task tag.</summary>
    public bool IsTask { get; set; }

    /// <summary>Whether the tag is checkable.</summary>
    public bool IsCheckable { get; set; }

    /// <summary>Whether the tag is completed.</summary>
    public bool IsCompleted { get; set; }

    /// <summary>Whether the tag is disabled.</summary>
    public bool IsDisabled { get; set; }

    /// <summary>Whether a task tag has not yet synchronized.</summary>
    public bool IsUnsynchronized { get; set; }

    /// <summary>Whether the task associated with this tag was removed.</summary>
    public bool IsRemoved { get; set; }

    /// <summary>Optional task due date.</summary>
    public DateTime? DueUtc { get; set; }

    /// <summary>Creation timestamp.</summary>
    public DateTime? CreatedUtc { get; set; }

    /// <summary>Completion timestamp.</summary>
    public DateTime? CompletedUtc { get; set; }

    /// <summary>Optional tag text color encoded as ARGB.</summary>
    public uint? TextColorArgb { get; set; }

    /// <summary>Optional tag highlight color encoded as ARGB.</summary>
    public uint? HighlightColorArgb { get; set; }
}
