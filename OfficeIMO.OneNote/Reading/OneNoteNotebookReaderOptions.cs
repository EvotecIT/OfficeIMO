namespace OfficeIMO.OneNote;

/// <summary>Controls offline notebook and table-of-contents loading.</summary>
public sealed class OneNoteNotebookReaderOptions {
    /// <summary>Default maximum nested section-group depth.</summary>
    public const int DefaultMaxSectionGroupDepth = 32;

    /// <summary>Limits used when reading each <c>.onetoc2</c> and <c>.one</c> revision store.</summary>
    public OneNoteReaderOptions OneNoteOptions { get; set; } = new OneNoteReaderOptions();

    /// <summary>Loads each section's pages and assets instead of returning TOC metadata only.</summary>
    public bool LoadSectionContent { get; set; } = true;

    /// <summary>
    /// Continues loading other sections when one section is corrupt, encrypted, or uses an
    /// unsupported legacy variant. The failed section remains in the hierarchy and a diagnostic is emitted.
    /// </summary>
    public bool ContinueOnSectionError { get; set; } = true;

    /// <summary>Recursively loads nested section-group TOC files.</summary>
    public bool RecurseSectionGroups { get; set; } = true;

    /// <summary>Includes the special <c>OneNote_RecycleBin</c> section group.</summary>
    public bool IncludeRecycleBin { get; set; }

    /// <summary>Maximum nested section-group depth.</summary>
    public int MaxSectionGroupDepth { get; set; } = DefaultMaxSectionGroupDepth;

    /// <summary>Maximum total number of TOC entries across the notebook.</summary>
    public int MaxNotebookEntries { get; set; } = 10000;

    /// <summary>Maximum number of files stored in a <c>.onepkg</c> cabinet.</summary>
    public int MaxPackageEntries { get; set; } = 10000;

    /// <summary>Maximum total bytes after expanding a <c>.onepkg</c> cabinet.</summary>
    public long MaxPackageExpandedBytes { get; set; } = 1024L * 1024L * 1024L;

    /// <summary>Maximum expanded size of one <c>.onepkg</c> entry.</summary>
    public long MaxPackageEntryBytes { get; set; } = OneNoteReaderOptions.DefaultMaxInputBytes;
}
