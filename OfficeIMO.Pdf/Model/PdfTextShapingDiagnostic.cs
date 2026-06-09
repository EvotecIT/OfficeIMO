using System.Collections.ObjectModel;
using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes text that can be written only through OfficeIMO.Pdf's current simplified scalar text path.
/// </summary>
public sealed class PdfTextShapingDiagnostic {
    internal PdfTextShapingDiagnostic(string source, int index, int scalar, string script, string code, string message) {
        Source = source ?? string.Empty;
        Index = index;
        CodePoint = "U+" + scalar.ToString(scalar <= 0xFFFF ? "X4" : "X", CultureInfo.InvariantCulture);
        Text = GetDisplayText(scalar);
        Script = script ?? string.Empty;
        Code = string.IsNullOrWhiteSpace(code) ? "unsupported-text-shaping" : code;
        Message = message ?? string.Empty;
    }

    /// <summary>Caller-provided source label such as a block, field, sheet, slide, or converter area.</summary>
    public string Source { get; }

    /// <summary>UTF-16 index of the first scalar requiring advanced text layout support.</summary>
    public int Index { get; }

    /// <summary>Unicode code point formatted as U+XXXX or U+XXXXX.</summary>
    public string CodePoint { get; }

    /// <summary>Scalar text that triggered the diagnostic.</summary>
    public string Text { get; }

    /// <summary>Script, feature, or layout family that requires advanced shaping or bidirectional layout.</summary>
    public string Script { get; }

    /// <summary>Stable warning code suitable for shared conversion reports.</summary>
    public string Code { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>
    /// Converts this shaping diagnostic to the shared conversion warning shape used by PDF adapters.
    /// </summary>
    /// <param name="converter">Converter or adapter name to place on the warning.</param>
    /// <returns>A shared conversion warning carrying this diagnostic and stable details.</returns>
    public PdfConversionWarning ToConversionWarning(string converter = "OfficeIMO.Pdf") {
        var details = new Dictionary<string, string> {
            ["index"] = Index.ToString(CultureInfo.InvariantCulture),
            ["codePoint"] = CodePoint,
            ["text"] = Text,
            ["script"] = Script
        };

        var layoutDiagnostic = new PdfLayoutDiagnostic(
            PdfLayoutDiagnosticKind.SimplifiedContent,
            Source,
            Message);

        return new PdfConversionWarning(
            converter,
            Code,
            Source,
            Message,
            PdfConversionWarningSeverity.Warning,
            layoutDiagnostic,
            new ReadOnlyDictionary<string, string>(details));
    }

    private static string GetDisplayText(int scalar) =>
        scalar < ' ' || scalar == '\u007F' || scalar > 0x10FFFF || (scalar >= 0xD800 && scalar <= 0xDFFF)
            ? string.Empty
            : char.ConvertFromUtf32(scalar);
}
