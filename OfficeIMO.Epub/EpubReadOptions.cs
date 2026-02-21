namespace OfficeIMO.Epub;

/// <summary>
/// Controls EPUB extraction behavior.
/// </summary>
public sealed class EpubReadOptions {
    /// <summary>
    /// Maximum number of chapter entries to emit.
    /// </summary>
    public int MaxChapters { get; set; } = 500;

    /// <summary>
    /// Optional maximum chapter size in bytes.
    /// </summary>
    public long? MaxChapterBytes { get; set; } = 4L * 1024 * 1024;

    /// <summary>
    /// When true, includes raw HTML per chapter.
    /// </summary>
    public bool IncludeRawHtml { get; set; }

    /// <summary>
    /// When true, chapter order is deterministic by internal path.
    /// </summary>
    public bool DeterministicOrder { get; set; } = true;

    /// <summary>
    /// When true, OPF spine ordering is preferred over archive path order.
    /// </summary>
    public bool PreferSpineOrder { get; set; } = true;

    /// <summary>
    /// When false, non-linear OPF spine items are skipped.
    /// </summary>
    public bool IncludeNonLinearSpineItems { get; set; } = true;

    /// <summary>
    /// When true, falls back to scanning XHTML/HTML entries when OPF/spine discovery fails.
    /// </summary>
    public bool FallbackToHtmlScan { get; set; } = true;
}
