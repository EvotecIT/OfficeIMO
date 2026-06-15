namespace OfficeIMO.Reader.Rtf;

/// <summary>
/// Options for RTF ingestion through the OfficeIMO.Reader adapter.
/// </summary>
public sealed class ReaderRtfOptions {
    /// <summary>
    /// Creates the default RTF reader profile.
    /// </summary>
    public static ReaderRtfOptions CreateOfficeIMOProfile() => new ReaderRtfOptions();

    /// <summary>
    /// Options passed to the dependency-free RTF parser and semantic binder.
    /// </summary>
    public RtfReadOptions? RtfReadOptions { get; set; }

    /// <summary>
    /// When true, emits one or more chunks per top-level RTF block. Default: true.
    /// </summary>
    public bool ChunkByBlock { get; set; } = true;

    /// <summary>
    /// When true, includes parsed header and footer text after body blocks.
    /// </summary>
    public bool IncludeHeadersAndFooters { get; set; } = true;

    /// <summary>
    /// When true, includes detached note text after body blocks.
    /// </summary>
    public bool IncludeNotes { get; set; } = true;

    /// <summary>
    /// When true, emits deterministic placeholders for images that appear as top-level blocks.
    /// </summary>
    public bool IncludeImagePlaceholders { get; set; } = true;

    /// <summary>
    /// When true, maps RTF parser and binder diagnostics into chunk warnings.
    /// </summary>
    public bool IncludeDiagnostics { get; set; } = true;

    /// <summary>
    /// Creates a defensive copy for handler registration reuse.
    /// </summary>
    public ReaderRtfOptions Clone() => new ReaderRtfOptions {
        RtfReadOptions = CloneReadOptions(RtfReadOptions),
        ChunkByBlock = ChunkByBlock,
        IncludeHeadersAndFooters = IncludeHeadersAndFooters,
        IncludeNotes = IncludeNotes,
        IncludeImagePlaceholders = IncludeImagePlaceholders,
        IncludeDiagnostics = IncludeDiagnostics
    };

    internal static RtfReadOptions? CloneReadOptions(RtfReadOptions? options) {
        if (options is null) return null;

        return new RtfReadOptions {
            MaxDepth = options.MaxDepth,
            WarnOnUnsupportedDestinations = options.WarnOnUnsupportedDestinations,
            WarnOnUnsupportedCodePages = options.WarnOnUnsupportedCodePages
        };
    }
}
