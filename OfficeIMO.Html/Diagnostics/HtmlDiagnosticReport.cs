namespace OfficeIMO.Html;

/// <summary>
/// Mutable report of diagnostics emitted by OfficeIMO HTML workflows.
/// </summary>
public sealed class HtmlDiagnosticReport : IReadOnlyList<HtmlDiagnostic> {
    private readonly List<HtmlDiagnostic> _diagnostics = new List<HtmlDiagnostic>();
    private readonly IReadOnlyList<HtmlDiagnostic> _readOnlyDiagnostics;

    /// <summary>Creates an empty diagnostic report.</summary>
    public HtmlDiagnosticReport() {
        _readOnlyDiagnostics = _diagnostics.AsReadOnly();
    }

    /// <summary>
    /// Diagnostics captured by the report in emission order.
    /// </summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics => _readOnlyDiagnostics;

    /// <summary>
    /// Number of diagnostics currently captured.
    /// </summary>
    public int Count => _diagnostics.Count;

    /// <summary>Gets a diagnostic by emission index.</summary>
    public HtmlDiagnostic this[int index] => _diagnostics[index];

    /// <summary>
    /// Indicates whether any captured diagnostic has error severity.
    /// </summary>
    public bool HasErrors => _diagnostics.Any(diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);

    /// <summary>
    /// Adds an existing diagnostic instance to the report.
    /// </summary>
    /// <param name="diagnostic">Diagnostic to add.</param>
    public void Add(HtmlDiagnostic diagnostic) {
        if (diagnostic == null) {
            throw new ArgumentNullException(nameof(diagnostic));
        }

        _diagnostics.Add(diagnostic);
    }

    /// <summary>
    /// Adds a diagnostic to the report.
    /// </summary>
    /// <param name="component">Component that emitted the diagnostic.</param>
    /// <param name="code">Stable diagnostic code.</param>
    /// <param name="message">Human-readable diagnostic message.</param>
    /// <param name="severity">Diagnostic severity.</param>
    /// <param name="source">Optional HTML, resource, or artifact source associated with the diagnostic.</param>
    /// <param name="detail">Optional low-level detail.</param>
    /// <param name="lossKind">Conversion fidelity impact represented by the diagnostic.</param>
    public void Add(
        string component,
        string code,
        string message,
        HtmlDiagnosticSeverity severity = HtmlDiagnosticSeverity.Warning,
        string? source = null,
        string? detail = null,
        HtmlConversionLossKind lossKind = HtmlConversionLossKind.None) {
        Add(new HtmlDiagnostic(component, code, message, severity, source, detail, lossKind));
    }

    /// <summary>Adds a diagnostic with typed source-to-target provenance.</summary>
    /// <param name="component">Component that emitted the diagnostic.</param>
    /// <param name="code">Stable diagnostic code.</param>
    /// <param name="message">Human-readable diagnostic message.</param>
    /// <param name="severity">Diagnostic severity.</param>
    /// <param name="source">Optional source.</param>
    /// <param name="detail">Optional low-level detail.</param>
    /// <param name="lossKind">Conversion fidelity impact.</param>
    /// <param name="sourceLocation">Optional typed HTML source location.</param>
    /// <param name="targetAddress">Optional target artifact address.</param>
    public void Add(
        string component,
        string code,
        string message,
        HtmlDiagnosticSeverity severity,
        string? source,
        string? detail,
        HtmlConversionLossKind lossKind,
        HtmlSemanticSourceLocation? sourceLocation,
        string? targetAddress) {
        Add(new HtmlDiagnostic(component, code, message, severity, source, detail, lossKind, sourceLocation, targetAddress));
    }

    /// <summary>
    /// Adds diagnostics to the report in enumeration order.
    /// </summary>
    /// <param name="diagnostics">Diagnostics to add.</param>
    public void AddRange(IEnumerable<HtmlDiagnostic> diagnostics) {
        if (diagnostics == null) {
            throw new ArgumentNullException(nameof(diagnostics));
        }

        foreach (HtmlDiagnostic diagnostic in diagnostics) {
            Add(diagnostic);
        }
    }

    /// <summary>
    /// Removes all diagnostics from the report.
    /// </summary>
    public void Clear() {
        _diagnostics.Clear();
    }

    /// <summary>
    /// Creates an independent copy of the current diagnostic report.
    /// </summary>
    /// <returns>A report containing the same immutable diagnostics.</returns>
    public HtmlDiagnosticReport Clone() {
        var clone = new HtmlDiagnosticReport();
        clone.AddRange(_diagnostics);
        return clone;
    }

    /// <inheritdoc />
    public IEnumerator<HtmlDiagnostic> GetEnumerator() => _diagnostics.GetEnumerator();

    /// <inheritdoc />
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
