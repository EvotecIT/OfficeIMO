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

    /// <summary>Optional width in OneNote half-inch layout units.</summary>
    public double? Width { get; set; }

    /// <summary>Optional height in OneNote half-inch layout units.</summary>
    public double? Height { get; set; }

    /// <summary>Named native page size, or automatic sizing when absent.</summary>
    public OneNotePageSize? PageSize { get; set; }

    /// <summary>Native page orientation.</summary>
    public OneNotePageOrientation? Orientation { get; set; }

    /// <summary>Page margins and their layout origins, in half-inch units.</summary>
    public OneNotePageMargins Margins { get; } = new OneNotePageMargins();

    /// <summary>Whether the page uses right-to-left root layout.</summary>
    public bool? RightToLeft { get; set; }

    /// <summary>Whether the page is marked read-only.</summary>
    public bool? IsReadOnly { get; set; }

    /// <summary>Whether OneNote resolves collisions between freely positioned outlines.</summary>
    public bool? ResolveChildCollisions { get; set; }

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

/// <summary>Native OneNote page-size identifiers.</summary>
public enum OneNotePageSize {
    /// <summary>The page automatically grows to fit its content.</summary>
    Automatic = 0,
    /// <summary>U.S. statement, 5.5 by 8.5 inches.</summary>
    Statement = 1,
    /// <summary>ANSI letter, 8.5 by 11 inches.</summary>
    Letter = 2,
    /// <summary>ANSI tabloid, 11 by 17 inches.</summary>
    Tabloid = 3,
    /// <summary>U.S. legal, 8.5 by 14 inches.</summary>
    Legal = 4,
    /// <summary>ISO A3.</summary>
    A3 = 5,
    /// <summary>ISO A4.</summary>
    A4 = 6,
    /// <summary>ISO A5.</summary>
    A5 = 7,
    /// <summary>ISO A6.</summary>
    A6 = 8,
    /// <summary>JIS B4.</summary>
    B4 = 9,
    /// <summary>JIS B5.</summary>
    B5 = 10,
    /// <summary>JIS B6.</summary>
    B6 = 11,
    /// <summary>Japanese postcard.</summary>
    JapanesePostcard = 12,
    /// <summary>Index card, 3 by 5 inches.</summary>
    IndexCard = 13,
    /// <summary>Billfold, 3.75 by 6.75 inches.</summary>
    Billfold = 14,
    /// <summary>Caller-defined width and height.</summary>
    Custom = 15
}

/// <summary>OneNote page orientation.</summary>
public enum OneNotePageOrientation {
    /// <summary>Portrait orientation.</summary>
    Portrait = 0,
    /// <summary>Landscape orientation.</summary>
    Landscape = 1
}

/// <summary>Page margins and margin origins in OneNote half-inch units.</summary>
public sealed class OneNotePageMargins {
    /// <summary>Left margin width.</summary>
    public double? Left { get; set; }
    /// <summary>Right margin width.</summary>
    public double? Right { get; set; }
    /// <summary>Top margin width.</summary>
    public double? Top { get; set; }
    /// <summary>Bottom margin width.</summary>
    public double? Bottom { get; set; }
    /// <summary>Horizontal margin origin.</summary>
    public double? OriginX { get; set; }
    /// <summary>Vertical margin origin.</summary>
    public double? OriginY { get; set; }
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

    /// <summary>MS-ONE action-item identity (0-99 for known normal tags, 100-105 for task tags; other native values are preserved).</summary>
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
