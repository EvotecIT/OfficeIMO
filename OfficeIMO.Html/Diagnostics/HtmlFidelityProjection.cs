namespace OfficeIMO.Html;

/// <summary>Projects compatibility-tolerant HTML diagnostics into the strict shared fidelity contract.</summary>
internal static class HtmlFidelityProjection {
    internal static OfficeConversionFidelityDiagnostic From(HtmlDiagnostic diagnostic) {
        if (diagnostic == null) throw new ArgumentNullException(nameof(diagnostic));
        string code = string.IsNullOrWhiteSpace(diagnostic.Code)
            ? "HTML_DIAGNOSTIC_UNCATEGORIZED"
            : diagnostic.Code;
        string source = string.IsNullOrWhiteSpace(diagnostic.Component)
            ? "OfficeIMO.Html"
            : diagnostic.Component;
        string message = string.IsNullOrWhiteSpace(diagnostic.Message)
            ? code + " was reported without a diagnostic message."
            : diagnostic.Message;
        return new OfficeConversionFidelityDiagnostic(
            code,
            message,
            diagnostic.LossKind,
            source,
            diagnostic.Provenance.SourceAddress);
    }
}
