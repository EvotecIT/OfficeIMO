namespace OfficeIMO.Html;

/// <summary>
/// Thrown when an HTML conversion cannot satisfy its requested semantic contract.
/// </summary>
public sealed class HtmlConversionException : InvalidOperationException {
    /// <summary>Creates an exception from structured conversion diagnostics.</summary>
    public HtmlConversionException(IEnumerable<HtmlDiagnostic> diagnostics)
        : this(ToDiagnosticList(diagnostics)) {
    }

    private HtmlConversionException(IReadOnlyList<HtmlDiagnostic> diagnostics)
        : base(CreateMessage(diagnostics)) {
        Diagnostics = diagnostics;
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
            ? "HTML conversion failed."
            : "HTML conversion failed: " + error.Message;
    }
}
