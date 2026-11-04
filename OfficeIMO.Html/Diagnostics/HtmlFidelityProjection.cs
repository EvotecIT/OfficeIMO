namespace OfficeIMO.Html;

/// <summary>Projects compatibility-tolerant HTML diagnostics into the strict shared fidelity contract.</summary>
internal static class HtmlFidelityProjection {
    internal static OfficeDiagnostic From(HtmlDiagnostic diagnostic) {
        ArgumentNullException.ThrowIfNull(diagnostic);
        string code = string.IsNullOrWhiteSpace(diagnostic.Code)
            ? "HTML_DIAGNOSTIC_UNCATEGORIZED"
            : diagnostic.Code;
        string source = string.IsNullOrWhiteSpace(diagnostic.Component)
            ? "OfficeIMO.Html"
            : diagnostic.Component;
        string message = string.IsNullOrWhiteSpace(diagnostic.Message)
            ? code + " was reported without a diagnostic message."
            : diagnostic.Message;
        return new global::OfficeIMO.OfficeDiagnostic(
            code,
            diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
            message,
            diagnostic.Provenance.SourceAddress,
            source,
            diagnostic.LossKind);
    }
}
