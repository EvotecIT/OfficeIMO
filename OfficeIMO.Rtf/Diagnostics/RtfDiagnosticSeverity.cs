namespace OfficeIMO.Rtf.Diagnostics;

/// <summary>
/// Severity assigned to an RTF parser, binder, or writer diagnostic.
/// </summary>
public enum RtfDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Info,
    /// <summary>Recoverable condition that may reduce fidelity.</summary>
    Warning,
    /// <summary>Invalid input or unsupported content that prevented a requested operation.</summary>
    Error
}
