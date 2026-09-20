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
        OfficeConversionLossKind? lossKind = null)
        : this(severity, code, message, source, lossKind, fidelitySource: null, fidelityLocation: null) {
    }

    internal OfficeImageExportDiagnostic(
        OfficeImageExportDiagnosticSeverity severity,
        string code,
        string message,
        string? source,
        OfficeConversionLossKind? lossKind,
        string? fidelitySource,
        string? fidelityLocation) {
        Severity = severity;
        Code = string.IsNullOrWhiteSpace(code) ? "ImageExportDiagnostic" : code;
        Message = message ?? string.Empty;
        Source = source;
        LossKind = severity == OfficeImageExportDiagnosticSeverity.Error
            && lossKind == OfficeConversionLossKind.None
                ? OfficeConversionLossKind.Failure
                : lossKind ?? InferLossKind(severity);
        FidelitySource = fidelitySource;
        FidelityLocation = fidelityLocation;
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
    public OfficeConversionLossKind LossKind { get; }

    internal string? FidelitySource { get; }

    internal string? FidelityLocation { get; }

    private static OfficeConversionLossKind InferLossKind(OfficeImageExportDiagnosticSeverity severity) => severity switch {
        OfficeImageExportDiagnosticSeverity.Warning => OfficeConversionLossKind.Approximation,
        OfficeImageExportDiagnosticSeverity.Error => OfficeConversionLossKind.Failure,
        _ => OfficeConversionLossKind.None
    };
}
