namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Markdown document plus explicit conversion diagnostics.</summary>
public sealed class AsciiDocMarkdownConversionResult {
    internal AsciiDocMarkdownConversionResult(MarkdownDoc document, IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        Document = document;
        Diagnostics = diagnostics;
    }

    /// <summary>Converted Markdown semantic document.</summary>
    public MarkdownDoc Document { get; }

    /// <summary>Loss, fallback, and omission diagnostics.</summary>
    public IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> Diagnostics { get; }

    /// <summary>True when at least one feature was simplified, source-fallbacked, or omitted.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.Outcome != AsciiDocMarkdownConversionOutcome.Converted);
}
