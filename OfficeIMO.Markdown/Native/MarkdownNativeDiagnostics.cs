namespace OfficeIMO.Markdown;

/// <summary>
/// Severity for native markdown projection diagnostics.
/// </summary>
public enum MarkdownNativeDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Info,

    /// <summary>Warning diagnostic.</summary>
    Warning,

    /// <summary>Error diagnostic.</summary>
    Error
}

/// <summary>
/// Diagnostic emitted while building the native markdown projection.
/// </summary>
public sealed class MarkdownNativeDiagnostic {
    internal MarkdownNativeDiagnostic(
        string id,
        string message,
        MarkdownNativeDiagnosticSeverity severity,
        MarkdownSourceSpan? sourceSpan = null,
        MarkdownNativeBlock? block = null,
        IReadOnlyList<MarkdownSourceSpan>? relatedSourceSpans = null) {
        Id = string.IsNullOrWhiteSpace(id) ? "native.diagnostic" : id.Trim();
        Message = message ?? string.Empty;
        Severity = severity;
        SourceSpan = sourceSpan;
        Block = block;
        RelatedSourceSpans = relatedSourceSpans ?? Array.Empty<MarkdownSourceSpan>();
    }

    /// <summary>Stable diagnostic identifier.</summary>
    public string Id { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Diagnostic severity.</summary>
    public MarkdownNativeDiagnosticSeverity Severity { get; }

    /// <summary>Source span associated with the diagnostic when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Native block associated with the diagnostic when available.</summary>
    public MarkdownNativeBlock? Block { get; }

    /// <summary>Additional source spans related to the diagnostic, such as individual transform input blocks.</summary>
    public IReadOnlyList<MarkdownSourceSpan> RelatedSourceSpans { get; }

    internal static MarkdownNativeDiagnostic FromTransform(MarkdownDocumentTransformDiagnostic diagnostic) {
        var sourceSpan = MarkdownTransformSourceSpanHelper.SelectMostSpecificSpan(
            diagnostic.AffectedFinalNodeSpan,
            diagnostic.AffectedOriginalNodeSpan,
            diagnostic.AffectedFinalBlockSpan,
            diagnostic.AffectedOriginalBlockSpan,
            diagnostic.AffectedSourceSpan);
        var message = string.IsNullOrWhiteSpace(diagnostic.TransformName)
            ? "Document transform ran while building the native markdown projection."
            : $"Document transform '{diagnostic.TransformName}' ran while building the native markdown projection.";

        return new MarkdownNativeDiagnostic(
            "native.transform",
            message,
            MarkdownNativeDiagnosticSeverity.Info,
            sourceSpan,
            relatedSourceSpans: diagnostic.AffectedSourceSpans);
    }
}
