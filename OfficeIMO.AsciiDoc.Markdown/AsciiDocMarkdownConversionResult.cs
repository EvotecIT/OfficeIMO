namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Markdown document plus explicit conversion diagnostics.</summary>
public sealed class AsciiDocToMarkdownResult {
    internal AsciiDocToMarkdownResult(MarkdownDoc value, IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }

    /// <summary>Converted Markdown semantic document.</summary>
    public MarkdownDoc Value { get; }

    /// <summary>Loss, fallback, and omission diagnostics.</summary>
    public IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> Diagnostics { get; }

    /// <summary>True when at least one feature was simplified, source-fallbacked, or omitted.</summary>
    public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.Outcome != AsciiDocMarkdownConversionOutcome.Converted);
}
