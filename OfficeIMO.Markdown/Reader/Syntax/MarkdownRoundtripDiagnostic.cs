namespace OfficeIMO.Markdown;

/// <summary>
/// Diagnostic emitted when a roundtrip writer cannot preserve source bytes losslessly.
/// </summary>
public sealed class MarkdownRoundtripDiagnostic {
    /// <summary>Creates a roundtrip diagnostic.</summary>
    public MarkdownRoundtripDiagnostic(
        string id,
        string message,
        MarkdownSourceSpan? sourceSpan = null,
        IReadOnlyList<MarkdownSourceSpan>? relatedSourceSpans = null) {
        Id = id ?? string.Empty;
        Message = message ?? string.Empty;
        SourceSpan = sourceSpan;
        RelatedSourceSpans = relatedSourceSpans ?? Array.Empty<MarkdownSourceSpan>();
    }

    /// <summary>Stable diagnostic identifier.</summary>
    public string Id { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Optional source span that caused or best explains the fallback.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Additional source spans related to the fallback, such as individual transform input blocks.</summary>
    public IReadOnlyList<MarkdownSourceSpan> RelatedSourceSpans { get; }
}
