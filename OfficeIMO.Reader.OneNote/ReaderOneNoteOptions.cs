using OfficeIMO.OneNote;

namespace OfficeIMO.Reader.OneNote;

/// <summary>Options for offline OneNote ingestion.</summary>
public sealed class ReaderOneNoteOptions {
    /// <summary>Native OneNote parser limits and compatibility settings.</summary>
    public OneNoteReaderOptions OneNoteOptions { get; set; } = new OneNoteReaderOptions();

    /// <summary>Notebook hierarchy and package-expansion settings for <c>.onetoc2</c> and <c>.onepkg</c> inputs.</summary>
    public OneNoteNotebookReaderOptions NotebookOptions { get; set; } = new OneNoteNotebookReaderOptions();

    /// <summary>
    /// Includes bounded image and embedded-file bytes in rich document results.
    /// Asset metadata is emitted regardless of this setting.
    /// </summary>
    public bool IncludeAssetPayloads { get; set; }

    /// <summary>Includes conflict-page snapshots in chunks, pages, assets, and links.</summary>
    public bool IncludeConflictPages { get; set; }

    /// <summary>Includes version-history snapshots in chunks, pages, assets, and links.</summary>
    public bool IncludeVersionHistory { get; set; }
}
