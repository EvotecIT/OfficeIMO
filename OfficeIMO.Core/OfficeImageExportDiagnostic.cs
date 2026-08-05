namespace OfficeIMO.Drawing;

/// <summary>
/// Structured diagnostic emitted while exporting Office content to an image.
/// </summary>
public sealed class OfficeImageExportDiagnostic {
    /// <summary>
    /// Creates a new image export diagnostic.
    /// </summary>
    public OfficeImageExportDiagnostic(
        OfficeImageExportDiagnosticSeverity severity,
        string code,
        string message,
        string? source = null,
        OfficeImageExportLossKind? lossKind = null) {
        Severity = severity;
        Code = string.IsNullOrWhiteSpace(code) ? "ImageExportDiagnostic" : code;
        Message = message ?? string.Empty;
        Source = source;
        LossKind = lossKind ?? InferLossKind(severity);
    }

    /// <summary>Diagnostic severity.</summary>
    public OfficeImageExportDiagnosticSeverity Severity { get; }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Optional source reference such as a cell range or sheet name.</summary>
    public string? Source { get; }

    /// <summary>Fidelity-loss classification used by aggregate reports and acceptance policies.</summary>
    public OfficeImageExportLossKind LossKind { get; }

    private static OfficeImageExportLossKind InferLossKind(OfficeImageExportDiagnosticSeverity severity) => severity switch {
        OfficeImageExportDiagnosticSeverity.Warning => OfficeImageExportLossKind.Approximation,
        OfficeImageExportDiagnosticSeverity.Error => OfficeImageExportLossKind.Failure,
        _ => OfficeImageExportLossKind.None
    };
}
