namespace OfficeIMO.Rtf;

/// <summary>
/// Options controlling RTF parsing and semantic binding.
/// </summary>
public sealed class RtfReadOptions {
    /// <summary>Default maximum nested group depth accepted by the parser.</summary>
    public const int DefaultMaxDepth = 512;

    /// <summary>Maximum nested group depth accepted by the semantic binder.</summary>
    public int MaxDepth { get; set; } = DefaultMaxDepth;

    /// <summary>Whether unsupported destinations should produce warning diagnostics.</summary>
    public bool WarnOnUnsupportedDestinations { get; set; } = true;

    /// <summary>Whether unsupported ANSI code pages should produce warning diagnostics before falling back to Windows-1252.</summary>
    public bool WarnOnUnsupportedCodePages { get; set; } = true;
}
