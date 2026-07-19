namespace OfficeIMO.Pdf;

/// <summary>Provides standalone text encoding and shaping preflight without creating a PDF document.</summary>
public static class PdfTextPreflight {
    /// <summary>Finds text that cannot be represented by the generated WinAnsi standard-font path.</summary>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeWinAnsi(
        string text,
        string source = "",
        string location = "") =>
        PdfTextDiagnostics.AnalyzeWinAnsiText(text, source, location);

    /// <summary>Finds text not covered by a supplied TrueType or OpenType/CFF font.</summary>
    public static IReadOnlyList<PdfTextEncodingDiagnostic> AnalyzeEmbeddedFont(
        string text,
        byte[] fontData,
        string source = "",
        string? fontName = null) =>
        PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontData, source, fontName);

    /// <summary>Finds right-to-left, complex-script shaping, mark-positioning, and line-breaking requirements.</summary>
    public static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedLayout(
        string text,
        string source = "") =>
        PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, source);

    /// <summary>Finds advanced-layout requirements with coverage evidence from a supplied font.</summary>
    public static IReadOnlyList<PdfTextShapingDiagnostic> AnalyzeAdvancedLayout(
        string text,
        byte[] fontData,
        string source = "",
        string? fontName = null) =>
        PdfTextDiagnostics.AnalyzeAdvancedTextLayout(text, fontData, source, fontName);
}
