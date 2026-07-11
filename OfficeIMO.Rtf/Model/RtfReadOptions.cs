namespace OfficeIMO.Rtf;

/// <summary>
/// Options controlling RTF parsing and semantic binding.
/// </summary>
public sealed class RtfReadOptions {
    /// <summary>Default maximum nested group depth accepted by the parser.</summary>
    public const int DefaultMaxDepth = 512;

    /// <summary>
    /// Creates the compatibility-oriented profile used by default. Only nesting depth is bounded;
    /// callers accepting untrusted content should use <see cref="CreateUntrustedProfile"/>.
    /// </summary>
    public static RtfReadOptions CreateOfficeIMOProfile() => new RtfReadOptions();

    /// <summary>
    /// Creates a conservative bounded profile for RTF received from an untrusted source.
    /// Embedded OLE objects and file-table references are omitted from the semantic model.
    /// </summary>
    public static RtfReadOptions CreateUntrustedProfile() => new RtfReadOptions {
        MaxDepth = 128,
        MaxInputBytes = 16L * 1024 * 1024,
        MaxInputCharacters = 16 * 1024 * 1024,
        MaxTokenCount = 1_000_000,
        MaxGroupCount = 250_000,
        MaxTextCharacters = 8_000_000,
        MaxBinaryBytesPerPayload = 4 * 1024 * 1024,
        MaxTotalBinaryBytes = 8L * 1024 * 1024,
        MaxImageCount = 256,
        MaxImageBytesPerImage = 4 * 1024 * 1024,
        MaxTotalImageBytes = 8L * 1024 * 1024,
        MaxObjectCount = 32,
        MaxObjectBytesPerObject = 4 * 1024 * 1024,
        MaxTotalObjectBytes = 8L * 1024 * 1024,
        MaxSemanticBlockCount = 100_000,
        ReadEmbeddedObjects = false,
        ReadFileReferences = false,
        HyperlinkPolicy = RtfHyperlinkReadPolicy.WebAndMailOnly
    };

    /// <summary>Maximum nested group depth accepted by the syntax parser and semantic binder.</summary>
    public int MaxDepth { get; set; } = DefaultMaxDepth;

    /// <summary>Maximum source bytes accepted by byte, file, and stream APIs.</summary>
    public long? MaxInputBytes { get; set; }

    /// <summary>Maximum source characters accepted before tokenization.</summary>
    public int? MaxInputCharacters { get; set; }

    /// <summary>Maximum number of tokenizer output tokens, including the end-of-file token.</summary>
    public int? MaxTokenCount { get; set; }

    /// <summary>Maximum number of source groups.</summary>
    public int? MaxGroupCount { get; set; }

    /// <summary>Maximum total number of source text characters represented by text tokens.</summary>
    public long? MaxTextCharacters { get; set; }

    /// <summary>Maximum bytes in one <c>\bin</c> payload.</summary>
    public int? MaxBinaryBytesPerPayload { get; set; }

    /// <summary>Maximum total bytes across <c>\bin</c> payloads.</summary>
    public long? MaxTotalBinaryBytes { get; set; }

    /// <summary>Maximum number of semantic images.</summary>
    public int? MaxImageCount { get; set; }

    /// <summary>Maximum decoded bytes in one semantic image.</summary>
    public int? MaxImageBytesPerImage { get; set; }

    /// <summary>Maximum decoded bytes across semantic images.</summary>
    public long? MaxTotalImageBytes { get; set; }

    /// <summary>Maximum number of semantic embedded or linked objects.</summary>
    public int? MaxObjectCount { get; set; }

    /// <summary>Maximum decoded bytes in one semantic object.</summary>
    public int? MaxObjectBytesPerObject { get; set; }

    /// <summary>Maximum decoded bytes across semantic objects.</summary>
    public long? MaxTotalObjectBytes { get; set; }

    /// <summary>Maximum number of semantic document blocks produced by binding.</summary>
    public int? MaxSemanticBlockCount { get; set; }

    /// <summary>Whether OLE object destinations are materialized in the semantic model.</summary>
    public bool ReadEmbeddedObjects { get; set; } = true;

    /// <summary>Whether file-table destinations are materialized in the semantic model.</summary>
    public bool ReadFileReferences { get; set; } = true;

    /// <summary>Controls hyperlink field targets materialized in the semantic model.</summary>
    public RtfHyperlinkReadPolicy HyperlinkPolicy { get; set; } = RtfHyperlinkReadPolicy.AllowAll;

    /// <summary>Whether unsupported destinations should produce warning diagnostics.</summary>
    public bool WarnOnUnsupportedDestinations { get; set; } = true;

    /// <summary>Whether unsupported ANSI code pages should produce warning diagnostics before falling back to Windows-1252.</summary>
    public bool WarnOnUnsupportedCodePages { get; set; } = true;
}
