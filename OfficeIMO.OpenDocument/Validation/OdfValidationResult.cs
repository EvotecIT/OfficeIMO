namespace OfficeIMO.OpenDocument;

/// <summary>Dependency-free structural validation result.</summary>
public sealed class OdfValidationResult {
    internal OdfValidationResult(IReadOnlyList<OdfDiagnostic> diagnostics) {
        Diagnostics = diagnostics;
    }

    /// <summary>Validation diagnostics.</summary>
    public IReadOnlyList<OdfDiagnostic> Diagnostics { get; }

    /// <summary>True when validation produced no error diagnostics.</summary>
    public bool IsValid => !Diagnostics.Any(diagnostic => diagnostic.Severity == OdfDiagnosticSeverity.Error);
}
