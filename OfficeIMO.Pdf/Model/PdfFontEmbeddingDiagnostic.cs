using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a font program that cannot be embedded through the current generated PDF font path.
/// </summary>
public sealed class PdfFontEmbeddingDiagnostic {
    internal PdfFontEmbeddingDiagnostic(string source, string fontName, string format, string code, string message) {
        Source = source ?? string.Empty;
        FontName = fontName ?? string.Empty;
        Format = format ?? string.Empty;
        Code = string.IsNullOrWhiteSpace(code) ? "unsupported-embedded-font" : code;
        Message = message ?? string.Empty;
        Severity = PdfConversionWarningSeverity.Error;
        LayoutDiagnosticKind = PdfLayoutDiagnosticKind.SkippedContent;
        Details = new ReadOnlyDictionary<string, string>(new Dictionary<string, string>());
    }

    internal PdfFontEmbeddingDiagnostic(
        string source,
        string fontName,
        string format,
        string code,
        string message,
        PdfConversionWarningSeverity severity,
        PdfLayoutDiagnosticKind layoutDiagnosticKind,
        IReadOnlyDictionary<string, string>? details = null) {
        Source = source ?? string.Empty;
        FontName = fontName ?? string.Empty;
        Format = format ?? string.Empty;
        Code = string.IsNullOrWhiteSpace(code) ? "unsupported-embedded-font" : code;
        Message = message ?? string.Empty;
        Severity = severity;
        LayoutDiagnosticKind = layoutDiagnosticKind;
        Details = new ReadOnlyDictionary<string, string>(CopyDetails(details));
    }

    /// <summary>Caller-provided source label such as a generated font slot, field, sheet, slide, or converter area.</summary>
    public string Source { get; }

    /// <summary>Configured PDF font name, when one was supplied.</summary>
    public string FontName { get; }

    /// <summary>Detected font container or outline format.</summary>
    public string Format { get; }

    /// <summary>Stable warning code suitable for shared conversion reports.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Severity assigned when this diagnostic is converted to a shared conversion warning.</summary>
    public PdfConversionWarningSeverity Severity { get; }

    /// <summary>Layout diagnostic kind assigned when this diagnostic is converted to a shared conversion warning.</summary>
    public PdfLayoutDiagnosticKind LayoutDiagnosticKind { get; }

    /// <summary>Additional stable diagnostic details.</summary>
    public IReadOnlyDictionary<string, string> Details { get; }

    /// <summary>
    /// Converts this font diagnostic to the shared conversion warning shape used by PDF adapters.
    /// </summary>
    /// <param name="converter">Converter or adapter name to place on the warning.</param>
    /// <returns>A shared conversion warning carrying this diagnostic and stable details.</returns>
    public PdfConversionWarning ToConversionWarning(string converter = "OfficeIMO.Pdf") {
        var details = new Dictionary<string, string> {
            ["fontName"] = FontName,
            ["format"] = Format
        };
        foreach (KeyValuePair<string, string> detail in Details) {
            details[detail.Key] = detail.Value;
        }

        var layoutDiagnostic = new PdfLayoutDiagnostic(
            LayoutDiagnosticKind,
            Source,
            Message);

        return new PdfConversionWarning(
            converter,
            Code,
            Source,
            Message,
            Severity,
            layoutDiagnostic,
            new ReadOnlyDictionary<string, string>(details));
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
