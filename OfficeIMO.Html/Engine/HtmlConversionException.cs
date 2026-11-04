namespace OfficeIMO.Html;

/// <summary>
/// Thrown when an HTML conversion cannot satisfy its requested semantic contract.
/// </summary>
public sealed class HtmlConversionException : OfficeConversionException {
    /// <summary>Creates an exception from structured conversion diagnostics.</summary>
    public HtmlConversionException(IEnumerable<HtmlDiagnostic> diagnostics)
        : this(ToDiagnosticList(diagnostics)) {
    }

    private HtmlConversionException(IReadOnlyList<HtmlDiagnostic> diagnostics)
        : this(new HtmlConversionReport(diagnostics)) {
    }

    internal HtmlConversionException(HtmlConversionReport report)
        : base(CreateMessage(report.NativeDiagnostics), report) {
        Diagnostics = report.NativeDiagnostics;
    }

    /// <summary>Diagnostics that caused the conversion to fail.</summary>
    public IReadOnlyList<HtmlDiagnostic> Diagnostics { get; }

    private static IReadOnlyList<HtmlDiagnostic> ToDiagnosticList(IEnumerable<HtmlDiagnostic> diagnostics) {
        if (diagnostics == null) {
            throw new ArgumentNullException(nameof(diagnostics));
        }

        return diagnostics.ToList().AsReadOnly();
    }

    private static string CreateMessage(IReadOnlyList<HtmlDiagnostic> diagnostics) {
        HtmlDiagnostic? error = diagnostics.FirstOrDefault(item => item.Severity == HtmlDiagnosticSeverity.Error);
        return error == null
            ? "HTML conversion did not satisfy the required no-loss contract."
                + string.Concat(diagnostics
                    .Where(item => item.LossKind != OfficeConversionLossKind.None
                        || item.Severity == HtmlDiagnosticSeverity.Warning)
                    .Select(item => " [" + item.Code + "] " + item.Message))
            : "HTML conversion failed: " + error.Message;
    }
}
