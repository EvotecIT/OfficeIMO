namespace OfficeIMO.Markdown;

/// <summary>
/// Result of a markdown roundtrip write attempt.
/// </summary>
public sealed class MarkdownRoundtripResult {
    /// <summary>Creates a roundtrip result.</summary>
    public MarkdownRoundtripResult(string markdown, IReadOnlyList<MarkdownRoundtripDiagnostic>? diagnostics = null) {
        Markdown = markdown ?? string.Empty;
        Diagnostics = diagnostics ?? Array.Empty<MarkdownRoundtripDiagnostic>();
    }

    /// <summary>The emitted markdown.</summary>
    public string Markdown { get; }

    /// <summary>Diagnostics describing any fallback from lossless source preservation.</summary>
    public IReadOnlyList<MarkdownRoundtripDiagnostic> Diagnostics { get; }

    /// <summary>True when the writer preserved the original source without fallbacks.</summary>
    public bool IsLossless => Diagnostics.Count == 0;
}
