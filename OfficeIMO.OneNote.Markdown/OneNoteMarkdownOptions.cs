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
    /// Resolves a Markdown destination for an image, attachment, recording, or ink payload.
    /// Returning <see langword="null"/> emits a readable placeholder without a link.
    /// </summary>
    public Func<OneNoteBinaryElement, string?>? AssetUriResolver { get; set; }

    internal OneNoteMarkdownOptions CloneValidated() {
        if (HeadingLevel < 1 || HeadingLevel > 6) throw new ArgumentOutOfRangeException(nameof(HeadingLevel), "HeadingLevel must be from 1 through 6.");
        return new OneNoteMarkdownOptions {
            IncludeConflictPages = IncludeConflictPages,
            IncludeVersionHistory = IncludeVersionHistory,
            HeadingLevel = HeadingLevel,
            AssetUriResolver = AssetUriResolver
        };
    }
}
