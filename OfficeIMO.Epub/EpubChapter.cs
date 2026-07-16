namespace OfficeIMO.Epub;

/// <summary>
/// Represents an extracted EPUB chapter-like content unit.
/// </summary>
public sealed class EpubChapter {
    /// <summary>
    /// 1-based chapter order.
    /// </summary>
    public int Order { get; internal set; }

    /// <summary>
    /// Internal EPUB entry path.
    /// </summary>
    public string Path { get; internal set; } = string.Empty;

    /// <summary>
    /// OPF manifest item id for this chapter when known.
    /// </summary>
    public string? ManifestId { get; internal set; }

    /// <summary>
    /// Media type from OPF manifest for this chapter when known.
    /// </summary>
    public string? MediaType { get; internal set; }

    /// <summary>
    /// 1-based spine index when chapter came from OPF spine ordering.
    /// </summary>
    public int? SpineIndex { get; internal set; }

    /// <summary>
    /// Indicates whether this chapter is marked linear in spine.
    /// </summary>
    public bool? IsLinear { get; internal set; }

    /// <summary>Effective declared layout for this spine item, when specified.</summary>
    public EpubRenditionLayout? RenditionLayout { get; internal set; }

    /// <summary>Whether this chapter is declared as pre-paginated fixed-layout content.</summary>
    public bool IsFixedLayout => RenditionLayout == EpubRenditionLayout.PrePaginated;

    /// <summary>Encryption declaration for this chapter resource, when present.</summary>
    public EpubEncryptionInfo? Encryption { get; internal set; }

    /// <summary>
    /// Best-effort chapter title.
    /// </summary>
    public string? Title { get; internal set; }

    /// <summary>
    /// Plain-text content extracted from chapter HTML.
    /// </summary>
    public string Text { get; internal set; } = string.Empty;

    /// <summary>
    /// Whether the chapter contains structured elements such as images, tables, forms, or media.
    /// </summary>
    public bool HasStructuredContent { get; internal set; }

    /// <summary>
    /// Optional raw HTML content.
    /// </summary>
    public string? Html { get; internal set; }
}
