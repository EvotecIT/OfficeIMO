namespace OfficeIMO.Html;

/// <summary>
/// Shared result contract for HTML conversions that produce a native target artifact.
/// </summary>
/// <typeparam name="T">Value produced by the conversion.</typeparam>
public abstract class HtmlConversionResult<T> {
    private readonly List<HtmlDiagnostic> _diagnostics = new List<HtmlDiagnostic>();
    private readonly IReadOnlyList<HtmlDiagnostic> _readOnlyDiagnostics;

    /// <summary>Creates a conversion result for the supplied value.</summary>
    protected HtmlConversionResult(T value) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        _readOnlyDiagnostics = _diagnostics.AsReadOnly();
        Report = new HtmlConversionReport(_readOnlyDiagnostics);
    }

    /// <summary>Value produced by the conversion.</summary>
    public T Value { get; }

    /// <summary>Diagnostics and fidelity outcome for this conversion.</summary>
    public HtmlConversionReport Report { get; }

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => Report.Succeeded;

    /// <summary>Whether the conversion approximated, omitted, or failed any source content.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>
    /// Returns the native artifact when conversion succeeded, or throws a structured conversion exception.
    /// </summary>
    public T RequireValue() {
        if (!Report.Succeeded) {
            throw new HtmlConversionException(Report.Diagnostics);
        }

        return Value;
    }

    /// <summary>
    /// Returns the native artifact only when conversion completed successfully and without approximation,
    /// omission, or failure diagnostics.
    /// </summary>
    public T RequireNoLoss() {
        T value = RequireValue();
        Report.RequireNoLoss();
        return value;
    }

    /// <summary>Adds one diagnostic while the result is being constructed.</summary>
    protected void AddDiagnostic(HtmlDiagnostic diagnostic) {
        _diagnostics.Add(diagnostic ?? throw new ArgumentNullException(nameof(diagnostic)));
    }

    /// <summary>Adds diagnostics while the result is being constructed.</summary>
    protected void AddDiagnostics(IEnumerable<HtmlDiagnostic> diagnostics) {
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));
        foreach (HtmlDiagnostic diagnostic in diagnostics) AddDiagnostic(diagnostic);
    }
}
