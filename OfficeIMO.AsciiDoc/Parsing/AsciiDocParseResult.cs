namespace OfficeIMO.AsciiDoc;

/// <summary>Result of parsing AsciiDoc source.</summary>
public sealed class AsciiDocParseResult {
    internal AsciiDocParseResult(AsciiDocDocument document, IReadOnlyList<AsciiDocDiagnostic> diagnostics) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        Diagnostics = diagnostics ?? Array.Empty<AsciiDocDiagnostic>();
    }

    /// <summary>Best recoverable typed document.</summary>
    public AsciiDocDocument Document { get; }

    /// <summary>Parser and recovery diagnostics.</summary>
    public IReadOnlyList<AsciiDocDiagnostic> Diagnostics { get; }

    /// <summary>True when any error diagnostic was produced.</summary>
    public bool HasErrors => Diagnostics.Any(static diagnostic => diagnostic.Severity == AsciiDocDiagnosticSeverity.Error);

    /// <summary>True when the syntax tree covers and exactly retains every source character.</summary>
    public bool IsLossless => Document.SyntaxTree.IsLossless;
}
