namespace OfficeIMO.Html;

/// <summary>
/// Severity for shared OfficeIMO HTML diagnostics.
/// </summary>
public enum HtmlDiagnosticSeverity {
    /// <summary>
    /// Informational diagnostic that does not indicate content loss.
    /// </summary>
    Info,

    /// <summary>
    /// Warning diagnostic for skipped, blocked, or degraded content.
    /// </summary>
    Warning,

    /// <summary>
    /// Error diagnostic for content that could not be processed as requested.
    /// </summary>
    Error
}
