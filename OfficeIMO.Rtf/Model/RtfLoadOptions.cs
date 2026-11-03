namespace OfficeIMO.Rtf;

/// <summary>
/// Options controlling RTF parsing and semantic binding.
/// </summary>
public sealed class RtfLoadOptions {
    /// <summary>Retains binary source slices and defers picture/object byte materialization until explicit native access or save. Default: false.</summary>
    /// <remarks>Syntax, source-size, binary declaration and semantic payload limits still apply. This does not enable embedded objects.
    /// Sources declaring fromhtml with pict or objdata destinations are unsupported in this profile because their HTML baseline
    /// requires eager semantic binary serialization; the ordinary profile remains available.</remarks>
    public bool DeferBinaryPayloads { get; set; }
    /// <summary>
    /// Creates the bounded profile for RTF received from an untrusted source. It matches the defaults:
    /// embedded OLE objects and file-table references are omitted from the semantic model.
    /// </summary>
    public static RtfLoadOptions UntrustedDefaults => new RtfLoadOptions();

    /// <summary>
    /// Creates the former permissive compatibility profile. This profile has no size or count limits,
    /// materializes embedded objects and file references, and accepts every hyperlink scheme. Use it only
    /// for trusted inputs when preserving legacy behavior is more important than bounded ingestion.
    /// </summary>
    public static RtfLoadOptions CreateCompatibilityProfile() => new RtfLoadOptions {
        Limits = RtfReadLimits.CreateUnbounded(),
        ReadEmbeddedObjects = true,
        ReadFileReferences = true,
        HyperlinkPolicy = RtfHyperlinkReadPolicy.AllowAll
    };

    /// <summary>Resource budgets applied by the tokenizer, syntax parser and semantic binder.</summary>
    public RtfReadLimits Limits { get; set; } = new RtfReadLimits();

    /// <summary>Whether OLE object destinations are materialized in the semantic model.</summary>
    public bool ReadEmbeddedObjects { get; set; }

    /// <summary>Whether file-table destinations are materialized in the semantic model.</summary>
    public bool ReadFileReferences { get; set; }

    /// <summary>Controls hyperlink field targets materialized in the semantic model.</summary>
    public RtfHyperlinkReadPolicy HyperlinkPolicy { get; set; } = RtfHyperlinkReadPolicy.WebAndMailOnly;

    /// <summary>Whether unsupported destinations should produce warning diagnostics.</summary>
    public bool WarnOnUnsupportedDestinations { get; set; } = true;

    /// <summary>Whether unsupported ANSI code pages should produce warning diagnostics before falling back to Windows-1252.</summary>
    public bool WarnOnUnsupportedCodePages { get; set; } = true;
}

/// <summary>
/// Resource budgets for RTF reading. Limits must be positive; a nullable limit set to null is not enforced.
/// </summary>
public sealed class RtfReadLimits {
    /// <summary>Default maximum nested group depth accepted by the parser.</summary>
    public const int DefaultMaxDepth = 128;

    /// <summary>Maximum depth used by the explicit legacy compatibility profile.</summary>
    public const int CompatibilityMaxDepth = 512;

    internal static RtfReadLimits CreateUnbounded() => new RtfReadLimits {
        MaxDepth = CompatibilityMaxDepth,
        MaxInputBytes = null,
        MaxInputCharacters = null,
        MaxTokens = null,
        MaxGroups = null,
        MaxTextCharacters = null,
        MaxBinaryPayloadBytes = null,
        MaxTotalBinaryBytes = null,
        MaxImages = null,
        MaxImageBytes = null,
        MaxTotalImageBytes = null,
        MaxObjects = null,
        MaxObjectBytes = null,
        MaxTotalObjectBytes = null,
        MaxSemanticBlocks = null
    };

    /// <summary>Maximum nested group depth accepted by the syntax parser and semantic binder.</summary>
    public int MaxDepth { get; set; } = DefaultMaxDepth;

    /// <summary>Maximum source bytes accepted by byte, file, and stream APIs.</summary>
    public long? MaxInputBytes { get; set; } = 16L * 1024 * 1024;

    /// <summary>Maximum source characters accepted before tokenization.</summary>
    public int? MaxInputCharacters { get; set; } = 16 * 1024 * 1024;

    /// <summary>Maximum number of tokenizer output tokens, including the end-of-file token.</summary>
    public int? MaxTokens { get; set; } = 1_000_000;

    /// <summary>Maximum number of source groups.</summary>
    public int? MaxGroups { get; set; } = 250_000;

    /// <summary>Maximum total number of source text characters represented by text tokens.</summary>
    public long? MaxTextCharacters { get; set; } = 8_000_000;

    /// <summary>Maximum bytes in one <c>\bin</c> payload.</summary>
    public int? MaxBinaryPayloadBytes { get; set; } = 4 * 1024 * 1024;

    /// <summary>Maximum total bytes across <c>\bin</c> payloads.</summary>
    public long? MaxTotalBinaryBytes { get; set; } = 8L * 1024 * 1024;

    /// <summary>Maximum number of semantic images.</summary>
    public int? MaxImages { get; set; } = 256;

    /// <summary>Maximum decoded bytes in one semantic image.</summary>
    public int? MaxImageBytes { get; set; } = 4 * 1024 * 1024;

    /// <summary>Maximum decoded bytes across semantic images.</summary>
    public long? MaxTotalImageBytes { get; set; } = 8L * 1024 * 1024;

    /// <summary>Maximum number of semantic embedded or linked objects.</summary>
    public int? MaxObjects { get; set; } = 32;

    /// <summary>Maximum decoded bytes in one semantic object.</summary>
    public int? MaxObjectBytes { get; set; } = 4 * 1024 * 1024;

    /// <summary>Maximum decoded bytes across semantic objects.</summary>
    public long? MaxTotalObjectBytes { get; set; } = 8L * 1024 * 1024;

    /// <summary>Maximum number of semantic document blocks produced by binding.</summary>
    public int? MaxSemanticBlocks { get; set; } = 100_000;

    /// <summary>Returns an independent copy of these limits.</summary>
    public RtfReadLimits Clone() => (RtfReadLimits)MemberwiseClone();
}
