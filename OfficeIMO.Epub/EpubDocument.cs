namespace OfficeIMO.Epub;

/// <summary>
/// Represents extracted EPUB content.
/// </summary>
public sealed class EpubDocument {
    /// <summary>
    /// Best-effort document title.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Package identifier from OPF metadata when available.
    /// </summary>
    public string? Identifier { get; set; }

    /// <summary>
    /// Primary language from OPF metadata when available.
    /// </summary>
    public string? Language { get; set; }

    /// <summary>
    /// Creator/author from OPF metadata when available.
    /// </summary>
    public string? Creator { get; set; }

    /// <summary>
    /// Internal path to the OPF package document when discovered.
    /// </summary>
    public string? OpfPath { get; set; }

    /// <summary>
    /// Extracted chapters.
    /// </summary>
    public IReadOnlyList<EpubChapter> Chapters { get; set; } = Array.Empty<EpubChapter>();

    /// <summary>
    /// Non-fatal warnings encountered during extraction.
    /// </summary>
    public IReadOnlyList<string> Warnings { get; set; } = Array.Empty<string>();
}
