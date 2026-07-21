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
    /// <param name="lossKind">Conversion fidelity impact represented by the diagnostic.</param>
    public HtmlDiagnostic(
        string component,
        string code,
        string message,
        HtmlDiagnosticSeverity severity = HtmlDiagnosticSeverity.Warning,
        string? source = null,
        string? detail = null,
        HtmlConversionLossKind lossKind = HtmlConversionLossKind.None)
        : this(component, code, message, severity, source, detail, lossKind, null, null) {
    }

    /// <summary>Creates a shared diagnostic with typed source-to-target provenance.</summary>
    /// <param name="component">Component that emitted the diagnostic.</param>
    /// <param name="code">Stable diagnostic code.</param>
    /// <param name="message">Human-readable diagnostic message.</param>
    /// <param name="severity">Diagnostic severity.</param>
    /// <param name="source">Optional HTML, resource, or artifact source.</param>
    /// <param name="detail">Optional low-level detail.</param>
    /// <param name="lossKind">Conversion fidelity impact.</param>
    /// <param name="sourceLocation">Optional typed HTML source location.</param>
    /// <param name="targetAddress">Optional target artifact address.</param>
    public HtmlDiagnostic(
        string component,
        string code,
        string message,
        HtmlDiagnosticSeverity severity,
        string? source,
        string? detail,
        HtmlConversionLossKind lossKind,
        HtmlSemanticSourceLocation? sourceLocation,
        string? targetAddress) {
        Component = component ?? throw new ArgumentNullException(nameof(component));
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Severity = severity;
        Source = source;
        Detail = detail;
        LossKind = lossKind;
        string sourceAddress = sourceLocation?.Selector
            ?? (string.IsNullOrWhiteSpace(source) ? "html:document" : source!);
        Provenance = new HtmlDiagnosticProvenance(
            sourceAddress,
            sourceLocation?.Line ?? 0,
            sourceLocation?.Column ?? 0,
            string.IsNullOrWhiteSpace(targetAddress) ? component : targetAddress!);
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

    /// <summary>
    /// Describes whether this diagnostic represents approximation, omission, or complete conversion failure.
    /// </summary>
    public HtmlConversionLossKind LossKind { get; }

    /// <summary>Stable source-to-target provenance. This is always present, with document/component fallbacks.</summary>
    public HtmlDiagnosticProvenance Provenance { get; }

    /// <summary>Creates an equivalent diagnostic with a more precise source and target address.</summary>
    public HtmlDiagnostic WithProvenance(HtmlSemanticSourceLocation? sourceLocation, string targetAddress) =>
        new HtmlDiagnostic(Component, Code, Message, Severity, Source, Detail, LossKind, sourceLocation, targetAddress);

    /// <inheritdoc />
    public override string ToString() {
        string source = string.IsNullOrWhiteSpace(Source) ? string.Empty : $" [{Source}]";
        return $"{Component}:{Code}:{Severity}{source} -> {Provenance.TargetAddress} {Message}";
    }
}
