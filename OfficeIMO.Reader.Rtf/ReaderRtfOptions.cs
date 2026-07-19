namespace OfficeIMO.Reader.Rtf;

/// <summary>
/// Options for RTF ingestion through the OfficeIMO.Reader adapter.
/// </summary>
public sealed class ReaderRtfOptions {
    /// <summary>Creates bounded Reader RTF options.</summary>
    public ReaderRtfOptions() { }

    /// <summary>
    /// Creates the default RTF reader profile.
    /// </summary>
    public static ReaderRtfOptions CreateOfficeIMOProfile() => new ReaderRtfOptions();

    /// <summary>Creates an explicitly trusted compatibility profile without core resource ceilings.</summary>
    public static ReaderRtfOptions CreateTrustedProfile() => new ReaderRtfOptions {
        RtfReadOptions = OfficeIMO.Rtf.RtfReadOptions.CreateOfficeIMOProfile()
    };

    /// <summary>
    /// Options passed to the shared RTF parser and semantic binder.
    /// </summary>
    public RtfReadOptions? RtfReadOptions { get; set; } = OfficeIMO.Rtf.RtfReadOptions.CreateUntrustedProfile();

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
    /// When true, reconstructs page membership from explicit page and section breaks.
    /// Automatic text overflow is not calculated. Default: false.
    /// </summary>
    public bool IncludePageLocations { get; set; }

    internal RtfConversionReport Report { get; } = new RtfConversionReport();

    /// <summary>
    /// Creates a defensive copy for handler registration reuse.
    /// </summary>
    public ReaderRtfOptions Clone() => new ReaderRtfOptions {
        RtfReadOptions = CloneReadOptions(RtfReadOptions),
        ChunkByBlock = ChunkByBlock,
        IncludeHeadersAndFooters = IncludeHeadersAndFooters,
        IncludeNotes = IncludeNotes,
        IncludeImagePlaceholders = IncludeImagePlaceholders,
        IncludeDiagnostics = IncludeDiagnostics,
        IncludePageLocations = IncludePageLocations
    };

    internal static RtfReadOptions? CloneReadOptions(RtfReadOptions? options) {
        if (options is null) return OfficeIMO.Rtf.RtfReadOptions.CreateUntrustedProfile();

        return new RtfReadOptions {
            MaxDepth = options.MaxDepth,
            MaxInputBytes = options.MaxInputBytes,
            MaxInputCharacters = options.MaxInputCharacters,
            MaxTokenCount = options.MaxTokenCount,
            MaxGroupCount = options.MaxGroupCount,
            MaxTextCharacters = options.MaxTextCharacters,
            MaxBinaryBytesPerPayload = options.MaxBinaryBytesPerPayload,
            MaxTotalBinaryBytes = options.MaxTotalBinaryBytes,
            MaxImageCount = options.MaxImageCount,
            MaxImageBytesPerImage = options.MaxImageBytesPerImage,
            MaxTotalImageBytes = options.MaxTotalImageBytes,
            MaxObjectCount = options.MaxObjectCount,
            MaxObjectBytesPerObject = options.MaxObjectBytesPerObject,
            MaxTotalObjectBytes = options.MaxTotalObjectBytes,
            MaxSemanticBlockCount = options.MaxSemanticBlockCount,
            ReadEmbeddedObjects = options.ReadEmbeddedObjects,
            ReadFileReferences = options.ReadFileReferences,
            HyperlinkPolicy = options.HyperlinkPolicy,
            WarnOnUnsupportedDestinations = options.WarnOnUnsupportedDestinations,
            WarnOnUnsupportedCodePages = options.WarnOnUnsupportedCodePages
        };
    }
}
