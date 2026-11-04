namespace OfficeIMO.Html;

/// <summary>Freezes HTML stage evidence at a native conversion boundary.</summary>
internal static class HtmlConversionEvidence {
    internal static OfficeConversionReport CreateReport(string source, string target, IEnumerable<HtmlDiagnostic> diagnostics,
        TimeSpan elapsed = default, string? path = null, long itemsWritten = 0) =>
        new OfficeConversionReport(source, target, path, itemsWritten, elapsed, diagnostics.Select(d => new OfficeDiagnostic(
            d.Code, d.Severity switch {
                HtmlDiagnosticSeverity.Info => OfficeDiagnosticSeverity.Information,
                HtmlDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                _ => OfficeDiagnosticSeverity.Warning
            }, d.Message, d.Source ?? d.Provenance.SourceAddress, d.Component, d.LossKind)));
}
