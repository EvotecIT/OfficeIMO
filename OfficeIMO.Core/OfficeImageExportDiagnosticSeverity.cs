namespace OfficeIMO.Drawing;

/// <summary>
/// Severity for image export diagnostics.
/// </summary>
public enum OfficeImageExportDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Info,

    /// <summary>Export completed with an approximation or skipped optional feature.</summary>
    Warning,

    /// <summary>Export could not fully complete the requested operation.</summary>
    Error
}
