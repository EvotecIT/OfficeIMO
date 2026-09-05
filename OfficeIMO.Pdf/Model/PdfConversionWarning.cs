using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

/// <summary>
/// Shared converter warning used by OfficeIMO PDF adapters so wrappers can consume one diagnostic shape.
/// </summary>
public sealed class PdfConversionWarning {
    /// <summary>
    /// Creates a converter warning.
    /// </summary>
    public PdfConversionWarning(
        string converter,
        string code,
        string source,
        string message,
        PdfConversionWarningSeverity severity = PdfConversionWarningSeverity.Warning,
        PdfLayoutDiagnostic? layoutDiagnostic = null,
        IReadOnlyDictionary<string, string>? details = null)
        : this(converter, code, source, message, severity,
            severity == PdfConversionWarningSeverity.Error ? OfficeConversionLossKind.Failure
                : severity == PdfConversionWarningSeverity.Information ? OfficeConversionLossKind.None : OfficeConversionLossKind.Approximation,
            layoutDiagnostic, details) {
    }

    /// <summary>
    /// Creates a warning with explicit fidelity impact, preserving loss-bearing
    /// informational diagnostics from upstream renderers and format adapters.
    /// </summary>
    public PdfConversionWarning(
        string converter,
        string code,
        string source,
        string message,
        PdfConversionWarningSeverity severity,
        OfficeConversionLossKind lossKind,
        PdfLayoutDiagnostic? layoutDiagnostic = null,
        IReadOnlyDictionary<string, string>? details = null) {
        if (lossKind < OfficeConversionLossKind.None || lossKind > OfficeConversionLossKind.Failure)
            throw new ArgumentOutOfRangeException(nameof(lossKind));
        Converter = string.IsNullOrWhiteSpace(converter) ? "OfficeIMO.Pdf" : converter;
        Code = string.IsNullOrWhiteSpace(code) ? "PdfConversionWarning" : code;
        Source = source ?? string.Empty;
        Message = message ?? string.Empty;
        Severity = severity;
        LossKind = lossKind == OfficeConversionLossKind.None && severity == PdfConversionWarningSeverity.Error
            ? OfficeConversionLossKind.Failure : lossKind;
        LayoutDiagnostic = layoutDiagnostic;
        Details = new ReadOnlyDictionary<string, string>(CopyDetails(details));
    }

    /// <summary>Converter or adapter that produced the warning.</summary>
    public string Converter { get; }

    /// <summary>Stable warning code suitable for assertions and wrapper routing.</summary>
    public string Code { get; }

    /// <summary>Source document area, page, sheet, slide, block type, or feature that produced the warning.</summary>
    public string Source { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Severity assigned by the converter.</summary>
    public PdfConversionWarningSeverity Severity { get; }

    /// <summary>Typed fidelity impact, independent of the diagnostic's presentation severity.</summary>
    public OfficeConversionLossKind LossKind { get; }

    /// <summary>Optional shared layout diagnostic when the warning maps to visible PDF layout or clipping behavior.</summary>
    public PdfLayoutDiagnostic? LayoutDiagnostic { get; }

    /// <summary>Additional converter-specific details such as sheet name, slide number, or feature name.</summary>
    public IReadOnlyDictionary<string, string> Details { get; }

    /// <inheritdoc />
    public override string ToString() {
        string prefix = string.IsNullOrWhiteSpace(Source)
            ? Converter + "/" + Code
            : Converter + "/" + Code + " [" + Source + "]";
        return prefix + ": " + Message;
    }

    private static Dictionary<string, string> CopyDetails(IReadOnlyDictionary<string, string>? details) {
        var copy = new Dictionary<string, string>();
        if (details == null) {
            return copy;
        }

        foreach (KeyValuePair<string, string> detail in details) {
            copy[detail.Key] = detail.Value;
        }

        return copy;
    }
}
