namespace OfficeIMO.Html;

/// <summary>
/// Immutable view of the diagnostics and fidelity outcome from one HTML conversion.
/// </summary>
public sealed class HtmlConversionReport {
    private readonly IReadOnlyList<HtmlDiagnostic> _diagnostics;

    internal HtmlConversionReport(IReadOnlyList<HtmlDiagnostic> diagnostics) {
        _diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
    }

    /// <summary>Structured conversion diagnostics in emission order.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Whether conversion completed without an error diagnostic.</summary>
    public bool Succeeded => !_diagnostics.Any(static diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);

    /// <summary>Whether the conversion approximated, omitted, or failed any source content.</summary>
    public bool HasLoss => _diagnostics.Any(static diagnostic => diagnostic.LossKind != HtmlConversionLossKind.None);

    /// <summary>Throws when the conversion failed or did not preserve the source faithfully.</summary>
    public void RequireNoLoss() {
        if (!Succeeded || HasLoss) {
            throw new HtmlConversionException(_diagnostics);
        }
    }
}
