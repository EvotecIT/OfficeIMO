namespace OfficeIMO.Html;

/// <summary>
/// Severity for an HTML/RTF conversion diagnostic.
/// </summary>
public enum HtmlRtfConversionDiagnosticSeverity {
    /// <summary>
    /// Informational diagnostic that does not indicate content loss.
    /// </summary>
    Info,

    /// <summary>
    /// Warning diagnostic for skipped or degraded content.
    /// </summary>
    Warning,

    /// <summary>
    /// Error diagnostic for content that could not be converted as requested.
    /// </summary>
    Error
}

/// <summary>
/// Describes skipped, degraded, or otherwise notable content observed during HTML/RTF conversion.
/// </summary>
public sealed class HtmlRtfConversionDiagnostic {
    /// <summary>
    /// Creates a conversion diagnostic.
    /// </summary>
    /// <param name="code">Stable diagnostic code.</param>
    /// <param name="message">Human-readable message.</param>
    /// <param name="severity">Diagnostic severity.</param>
    /// <param name="source">Optional HTML/resource source associated with the diagnostic.</param>
    /// <param name="detail">Optional low-level detail, such as an exception type or status text.</param>
    public HtmlRtfConversionDiagnostic(string code, string message, HtmlRtfConversionDiagnosticSeverity severity = HtmlRtfConversionDiagnosticSeverity.Warning, string? source = null, string? detail = null) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Severity = severity;
        Source = source;
        Detail = detail;
    }

    /// <summary>
    /// Stable diagnostic code.
    /// </summary>
    public string Code { get; }

    /// <summary>
    /// Human-readable diagnostic message.
    /// </summary>
    public string Message { get; }

    /// <summary>
    /// Diagnostic severity.
    /// </summary>
    public HtmlRtfConversionDiagnosticSeverity Severity { get; }

    /// <summary>
    /// Optional HTML/resource source associated with the diagnostic.
    /// </summary>
    public string? Source { get; }

    /// <summary>
    /// Optional low-level detail, such as an exception type or status text.
    /// </summary>
    public string? Detail { get; }
}
