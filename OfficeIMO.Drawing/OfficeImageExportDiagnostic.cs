namespace OfficeIMO.Drawing;

/// <summary>
/// Structured diagnostic emitted while exporting Office content to an image.
/// </summary>
public sealed class OfficeImageExportDiagnostic {
    /// <summary>
    /// Creates a new image export diagnostic.
    /// </summary>
    public OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity severity, string code, string message, string? source = null) {
        Severity = severity;
        Code = string.IsNullOrWhiteSpace(code) ? "ImageExportDiagnostic" : code;
        Message = message ?? string.Empty;
        Source = source;
    }

    /// <summary>Diagnostic severity.</summary>
    public OfficeImageExportDiagnosticSeverity Severity { get; }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Optional source reference such as a cell range or sheet name.</summary>
    public string? Source { get; }
}
