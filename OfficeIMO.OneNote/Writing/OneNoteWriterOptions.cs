namespace OfficeIMO.OneNote;

/// <summary>Controls native offline OneNote serialization.</summary>
public sealed class OneNoteWriterOptions {
    /// <summary>Hard maximum accepted for recursive writer traversal limits.</summary>
    public const int MaximumTraversalDepth = 256;

    /// <summary>Default maximum nesting depth for serializable content elements.</summary>
    public const int DefaultMaxContentDepth = OneNoteReaderOptions.DefaultMaxPropertySetDepth;

    /// <summary>Default maximum nesting depth for notebook section groups.</summary>
    public const int DefaultMaxSectionGroupDepth = OneNoteNotebookReaderOptions.DefaultMaxSectionGroupDepth;

    /// <summary>
    /// Physical encoding for generated <c>.one</c> and <c>.onetoc2</c> payloads.
    /// <see cref="OneNoteStorageFormat.Unknown"/> preserves a loaded artifact's source encoding
    /// and defaults new artifacts to <see cref="OneNoteStorageFormat.RevisionStore"/>.
    /// </summary>
    public OneNoteStorageFormat StorageFormat { get; set; } = OneNoteStorageFormat.Unknown;

    /// <summary>Maximum serialized output size. The default is 512 MiB.</summary>
    public long MaxOutputBytes { get; set; } = OneNoteReaderOptions.DefaultMaxInputBytes;

    /// <summary>
    /// Reads the generated artifact back and verifies page identity, order, relationships, structural content,
    /// rich-text runs, supported layout/media metadata, and binary payload resolution state before returning it.
    /// </summary>
    public bool ValidateRoundTrip { get; set; } = true;

    /// <summary>Maximum number of files emitted into a notebook directory or <c>.onepkg</c> archive.</summary>
    public int MaxPackageEntries { get; set; } = 10000;

    /// <summary>
    /// Maximum nesting depth for notebook section groups.
    /// Values must be from 1 through <see cref="MaximumTraversalDepth"/>.
    /// </summary>
    public int MaxSectionGroupDepth { get; set; } = DefaultMaxSectionGroupDepth;

    /// <summary>
    /// Maximum nesting depth across conflict and version-history page relationships.
    /// Values must be from 1 through <see cref="MaximumTraversalDepth"/>.
    /// </summary>
    public int MaxPageRelationshipDepth { get; set; } = OneNoteReaderOptions.DefaultMaxPageRelationshipDepth;

    /// <summary>
    /// Maximum recursive nesting depth for outlines, paragraphs, and table-cell content.
    /// Values must be from 1 through <see cref="MaximumTraversalDepth"/>.
    /// </summary>
    public int MaxContentDepth { get; set; } = DefaultMaxContentDepth;

    /// <summary>
    /// Preserves source objects, properties, and relationships that are not replaced by typed edits.
    /// This is enabled by default for sections loaded by OfficeIMO.
    /// </summary>
    public bool PreserveUnknownData { get; set; } = true;
}
