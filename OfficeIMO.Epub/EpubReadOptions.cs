namespace OfficeIMO.Epub;

/// <summary>
/// Controls EPUB extraction behavior.
/// </summary>
public sealed class EpubReadOptions {
    /// <summary>Maximum compressed EPUB package size read from a file or stream.</summary>
    public long MaxPackageBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Maximum number of ZIP entries indexed from the container.</summary>
    public int MaxArchiveEntries { get; set; } = 10_000;

    /// <summary>Maximum combined uncompressed size declared by ZIP entries.</summary>
    public long MaxTotalUncompressedBytes { get; set; } = 4L * 1024 * 1024 * 1024;

    /// <summary>Maximum size of container, OPF, navigation, NCX, or encryption XML read into memory.</summary>
    public long MaxPackageMetadataBytes { get; set; } = 4L * 1024 * 1024;

    /// <summary>Maximum number of ordered OPF metadata declarations retained.</summary>
    public int MaxMetadataItems { get; set; } = 4_096;

    /// <summary>Maximum combined number of table-of-contents, page-list, and landmark items retained.</summary>
    public int MaxNavigationItems { get; set; } = 10_000;

    /// <summary>Maximum retained navigation nesting depth.</summary>
    public int MaxNavigationDepth { get; set; } = 64;

    /// <summary>
    /// Maximum number of chapter entries to emit.
    /// </summary>
    public int MaxChapters { get; set; } = 500;

    /// <summary>
    /// Optional maximum chapter size in bytes.
    /// </summary>
    public long? MaxChapterBytes { get; set; } = 4L * 1024 * 1024;

    /// <summary>
    /// Maximum combined uncompressed bytes of raw chapter HTML retained in memory.
    /// </summary>
    public long MaxTotalRawHtmlBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>
    /// When true, includes raw HTML per chapter.
    /// </summary>
    public bool IncludeRawHtml { get; set; }

    /// <summary>
    /// When true, includes bounded manifest resource payloads.
    /// </summary>
    public bool IncludeResourceData { get; set; }

    /// <summary>Maximum number of manifest resources returned.</summary>
    public int MaxResources { get; set; } = 2_000;

    /// <summary>Maximum payload size for one manifest resource.</summary>
    public long MaxResourceBytes { get; set; } = 8L * 1024 * 1024;

    /// <summary>Maximum combined payload size returned for manifest resources.</summary>
    public long MaxTotalResourceBytes { get; set; } = 64L * 1024 * 1024;

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
