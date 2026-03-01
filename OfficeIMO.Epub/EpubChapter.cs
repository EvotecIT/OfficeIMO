namespace OfficeIMO.Epub;

/// <summary>
/// Represents an extracted EPUB chapter-like content unit.
/// </summary>
public sealed class EpubChapter {
    /// <summary>
    /// 1-based chapter order.
    /// </summary>
    public int Order { get; set; }

    /// <summary>
    /// Internal EPUB entry path.
    /// </summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>
    /// OPF manifest item id for this chapter when known.
    /// </summary>
    public string? ManifestId { get; set; }

    /// <summary>
    /// Media type from OPF manifest for this chapter when known.
    /// </summary>
    public string? MediaType { get; set; }

    /// <summary>
    /// 1-based spine index when chapter came from OPF spine ordering.
    /// </summary>
    public int? SpineIndex { get; set; }

    /// <summary>
    /// Indicates whether this chapter is marked linear in spine.
    /// </summary>
    public bool? IsLinear { get; set; }

    /// <summary>
    /// Best-effort chapter title.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Plain-text content extracted from chapter HTML.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional raw HTML content.
    /// </summary>
    public string? Html { get; set; }
}
