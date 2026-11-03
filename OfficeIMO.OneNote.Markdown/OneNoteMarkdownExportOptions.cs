namespace OfficeIMO.OneNote.Markdown;

/// <summary>Controls semantic projection of offline OneNote models to Markdown.</summary>
public sealed class OneNoteMarkdownOptions {
    /// <summary>Includes conflict copies after their current page.</summary>
    public bool IncludeConflictPages { get; set; }

    /// <summary>Includes version-history snapshots after their current page.</summary>
    public bool IncludeVersionHistory { get; set; }

    /// <summary>First heading level used by a section or notebook title.</summary>
    public int HeadingLevel { get; set; } = 1;

    /// <summary>
    /// Maximum nested section-group depth projected from a notebook.
    /// Values must be from 1 through <see cref="OneNoteWriterOptions.MaximumTraversalDepth"/>.
    /// </summary>
    public int MaxSectionGroupDepth { get; set; } = OneNoteNotebookReaderOptions.DefaultMaxSectionGroupDepth;

    /// <summary>
    /// Maximum nesting depth across included conflict and version-history page relationships.
    /// Values must be from 1 through <see cref="OneNoteWriterOptions.MaximumTraversalDepth"/>.
    /// </summary>
    public int MaxPageRelationshipDepth { get; set; } = OneNoteReaderOptions.DefaultMaxPageRelationshipDepth;

    /// <summary>
    /// Maximum recursive nesting depth for outlines, paragraphs, and table-cell content.
    /// Values must be from 1 through <see cref="OneNoteWriterOptions.MaximumTraversalDepth"/>.
    /// </summary>
    public int MaxContentDepth { get; set; } = OneNoteWriterOptions.DefaultMaxContentDepth;

    /// <summary>
    /// Resolves a Markdown destination for an image, attachment, recording, or ink payload.
    /// Returning <see langword="null"/> emits a readable placeholder without a link.
    /// </summary>
    public Func<OneNoteBinaryElement, string?>? AssetUriResolver { get; set; }

    /// <summary>Creates a validated independent snapshot for one projection operation.</summary>
    public OneNoteMarkdownOptions Clone() => CloneValidated();

    internal OneNoteMarkdownOptions CloneValidated() {
        if (HeadingLevel < 1 || HeadingLevel > 6) throw new ArgumentOutOfRangeException(nameof(HeadingLevel), "HeadingLevel must be from 1 through 6.");
        ValidateDepth(MaxSectionGroupDepth, nameof(MaxSectionGroupDepth));
        ValidateDepth(MaxPageRelationshipDepth, nameof(MaxPageRelationshipDepth));
        ValidateDepth(MaxContentDepth, nameof(MaxContentDepth));
        return new OneNoteMarkdownOptions {
            IncludeConflictPages = IncludeConflictPages,
            IncludeVersionHistory = IncludeVersionHistory,
            HeadingLevel = HeadingLevel,
            MaxSectionGroupDepth = MaxSectionGroupDepth,
            MaxPageRelationshipDepth = MaxPageRelationshipDepth,
            MaxContentDepth = MaxContentDepth,
            AssetUriResolver = AssetUriResolver
        };
    }

    private static void ValidateDepth(int value, string parameterName) {
        if (value < 1 || value > OneNoteWriterOptions.MaximumTraversalDepth) {
            throw new ArgumentOutOfRangeException(
                parameterName,
                "Projection traversal depths must be between 1 and " + OneNoteWriterOptions.MaximumTraversalDepth + ".");
        }
    }
}
