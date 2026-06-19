namespace OfficeIMO.Html;

/// <summary>
/// Shared diagnostic record for OfficeIMO HTML ingestion, conversion, and artifact workflows.
/// </summary>
public sealed class HtmlDiagnostic {
    /// <summary>
    /// Creates a shared HTML diagnostic.
    /// </summary>
    /// <param name="component">Component that emitted the diagnostic, such as <c>OfficeIMO.Word.Html</c>.</param>
    /// <param name="code">Stable diagnostic code.</param>
    /// <param name="message">Human-readable diagnostic message.</param>
    /// <param name="severity">Diagnostic severity.</param>
    /// <param name="source">Optional HTML, resource, or artifact source associated with the diagnostic.</param>
    /// <param name="detail">Optional low-level detail, such as an exception type, status text, or limit data.</param>
    public HtmlDiagnostic(string component, string code, string message, HtmlDiagnosticSeverity severity = HtmlDiagnosticSeverity.Warning, string? source = null, string? detail = null) {
        Component = component ?? throw new ArgumentNullException(nameof(component));
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Severity = severity;
        Source = source;
        Detail = detail;
    }

    /// <summary>
    /// Component that emitted the diagnostic.
    /// </summary>
    public string Component { get; }

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
    public HtmlDiagnosticSeverity Severity { get; }

    /// <summary>
    /// Optional HTML, resource, or artifact source associated with the diagnostic.
    /// </summary>
    public string? Source { get; }

    /// <summary>
    /// Optional low-level detail, such as an exception type, status text, or limit data.
    /// </summary>
    public string? Detail { get; }

    /// <inheritdoc />
    public override string ToString() {
        string source = string.IsNullOrWhiteSpace(Source) ? string.Empty : $" [{Source}]";
        return $"{Component}:{Code}:{Severity}{source} {Message}";
    }
}
