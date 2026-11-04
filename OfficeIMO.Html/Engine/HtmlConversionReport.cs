namespace OfficeIMO.Html;

/// <summary>
/// Immutable view of the diagnostics and fidelity outcome from one HTML conversion.
/// </summary>
public sealed class HtmlConversionReport : OfficeConversionReport {
    private readonly IReadOnlyList<HtmlDiagnostic> _diagnostics;

    internal HtmlConversionReport(IReadOnlyList<HtmlDiagnostic> diagnostics) {
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
    }

    /// <summary>Structured conversion diagnostics in emission order.</summary>
    public IReadOnlyList<HtmlDiagnostic> NativeDiagnostics => _diagnostics;

    /// <summary>Category-preserving conversion diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics =>
        Array.AsReadOnly(_diagnostics.Select(static diagnostic =>
            HtmlFidelityProjection.From(diagnostic)).ToArray());

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => !_diagnostics.Any(static diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);

    /// <summary>Whether the conversion approximated, omitted, or failed any source content.</summary>
    public override bool HasLoss => _diagnostics.Any(static diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Creates the exception that carries the HTML diagnostics of a lossy or failed conversion.</summary>
    protected override OfficeConversionException CreateLossException() => new HtmlConversionException(this);

}
