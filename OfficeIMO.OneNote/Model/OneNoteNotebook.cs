namespace OfficeIMO.OneNote;

/// <summary>
/// A OneNote notebook with ordered section groups and sections.
/// </summary>
public sealed class OneNoteNotebook {
    internal OneNoteExtendedGuid? TableOfContentsRootObjectId { get; set; }

    /// <summary>Notebook identity when available. Serialization assigns and retains an identity for a new notebook.</summary>
    public Guid? Id { get; set; }

    /// <summary>Notebook display name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Source folder or package path.</summary>
    public string? SourcePath { get; set; }

    /// <summary>Physical encoding of the loaded root <c>.onetoc2</c>, or unknown for a new notebook.</summary>
    public OneNoteStorageFormat TableOfContentsStorageFormat { get; internal set; }

    /// <summary>Notebook color encoded as ARGB, when present.</summary>
    public uint? ColorArgb { get; set; }

    /// <summary>Whether the notebook enables version history.</summary>
    public bool? HistoryEnabled { get; set; }

    /// <summary>Top-level section groups in notebook order.</summary>
    public IList<OneNoteSectionGroup> SectionGroups { get; } = new List<OneNoteSectionGroup>();

    /// <summary>Top-level sections in notebook order.</summary>
    public IList<OneNoteSection> Sections { get; } = new List<OneNoteSection>();

    /// <summary>Opaque notebook objects preserved for loss-aware writing.</summary>
    public IList<OneNoteOpaqueObject> UnknownObjects { get; } = new List<OneNoteOpaqueObject>();

    /// <summary>Diagnostics produced while loading the notebook.</summary>
    public IList<OneNoteDiagnostic> Diagnostics { get; } = new List<OneNoteDiagnostic>();
}

/// <summary>
/// A nested section-group folder in a OneNote notebook.
/// </summary>
public sealed class OneNoteSectionGroup {
    internal OneNoteExtendedGuid? TableOfContentsRootObjectId { get; set; }

    /// <summary>
    /// Relative navigation order among sibling sections and section groups when loaded from or written to a notebook TOC.
    /// Items without an explicit order follow ordered items in their collection order.
    /// </summary>
    public uint? TableOfContentsOrder { get; set; }

    /// <summary>Section-group identity when available. Serialization assigns and retains an identity for a new group.</summary>
    public Guid? Id { get; set; }

    /// <summary>Section-group display name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Relative folder path inside the notebook.</summary>
    public string? RelativePath { get; set; }

    /// <summary>Physical encoding of the loaded group <c>.onetoc2</c>, or unknown for a new group.</summary>
    public OneNoteStorageFormat TableOfContentsStorageFormat { get; internal set; }

    /// <summary>Whether this group represents OneNote's recycle bin.</summary>
    public bool IsRecycleBin { get; set; }

    /// <summary>Nested section groups in display order.</summary>
    public IList<OneNoteSectionGroup> SectionGroups { get; } = new List<OneNoteSectionGroup>();

    /// <summary>Sections in display order.</summary>
    public IList<OneNoteSection> Sections { get; } = new List<OneNoteSection>();

    /// <summary>Opaque section-group data preserved for loss-aware writing.</summary>
    public IList<OneNoteOpaqueObject> UnknownObjects { get; } = new List<OneNoteOpaqueObject>();
}

/// <summary>
/// A OneNote section and its ordered pages.
/// </summary>
public sealed class OneNoteSection {
    internal OneNoteSectionPreservationState? PreservationState { get; set; }

    /// <summary>
    /// Relative navigation order among sibling sections and section groups when loaded from or written to a notebook TOC.
    /// Items without an explicit order follow ordered items in their collection order.
    /// </summary>
    public uint? TableOfContentsOrder { get; set; }

    /// <summary>Section identity when available. Serialization assigns and retains an identity for a new section.</summary>
    public Guid? Id { get; set; }

    /// <summary>Section display name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Source <c>.one</c> path or package entry.</summary>
    public string? SourcePath { get; set; }

    /// <summary>Physical encoding of the loaded section, or unknown for a newly created section.</summary>
    public OneNoteStorageFormat StorageFormat { get; internal set; }

    /// <summary>Section color encoded as ARGB, when present.</summary>
    public uint? ColorArgb { get; set; }

    /// <summary>Ordered top-level pages and subpages.</summary>
    public IList<OneNotePage> Pages { get; } = new List<OneNotePage>();

    /// <summary>Revision metadata associated with the section object space.</summary>
    public IList<OneNoteRevision> Revisions { get; } = new List<OneNoteRevision>();

    /// <summary>Opaque section objects preserved for loss-aware writing.</summary>
    public IList<OneNoteOpaqueObject> UnknownObjects { get; } = new List<OneNoteOpaqueObject>();

    /// <summary>Diagnostics produced while loading the section.</summary>
    public IList<OneNoteDiagnostic> Diagnostics { get; } = new List<OneNoteDiagnostic>();

    /// <summary>Writes this section to an offline <c>.one</c> file.</summary>
    public void Save(string path, OneNoteWriterOptions? options = null) => OneNoteSectionWriter.Write(this, path, options);

    /// <summary>Serializes this section as an offline <c>.one</c> payload.</summary>
    public byte[] ToByteArray(OneNoteWriterOptions? options = null) => OneNoteSectionWriter.Write(this, options);
}

/// <summary>
/// Revision identity and provenance for an object space.
/// </summary>
public sealed class OneNoteRevision {
    /// <summary>Revision identifier.</summary>
    public OneNoteExtendedGuid? Id { get; set; }

    /// <summary>Base revision identifier, when declared.</summary>
    public OneNoteExtendedGuid? BaseRevisionId { get; set; }

    /// <summary>Revision creation time, when available.</summary>
    public DateTime? CreatedUtc { get; set; }

    /// <summary>Author associated with the revision.</summary>
    public string? Author { get; set; }

    /// <summary>Whether this revision contributes to the current materialized state.</summary>
    public bool IsCurrent { get; set; }

    /// <summary>Whether this revision represents version-history content.</summary>
    public bool IsVersionHistory { get; set; }

    /// <summary>Opaque revision data preserved for loss-aware writing.</summary>
    public IList<OneNoteOpaqueObject> UnknownObjects { get; } = new List<OneNoteOpaqueObject>();
}
