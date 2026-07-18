using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static class VisioImageExportFontDiagnostics {
    internal static void Append(
        VisioPage page,
        OfficeFontFaceCollection fonts,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        string source) {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (VisioShape shape in page.Shapes) {
            AppendShape(shape, fonts, diagnostics, seen, source);
        }
        foreach (VisioConnector connector in page.Connectors) {
            AppendText(connector.Label, connector.TextStyle, fonts, diagnostics, seen, source);
        }
    }

    private static void AppendShape(
        VisioShape shape,
        OfficeFontFaceCollection fonts,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        HashSet<string> seen,
        string source) {
        AppendText(shape.Text, shape.TextStyle, fonts, diagnostics, seen, source);
        foreach (VisioShape child in shape.Children) {
            AppendShape(child, fonts, diagnostics, seen, source);
        }
    }

    private static void AppendText(
        string? text,
        VisioTextStyle? style,
        OfficeFontFaceCollection fonts,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        HashSet<string> seen,
        string source) {
        if (string.IsNullOrEmpty(text)) return;
        string family = string.IsNullOrWhiteSpace(style?.FontFamily)
            ? "Aptos, Calibri, Arial, sans-serif"
            : style!.FontFamily!;
        OfficeFontStyle fontStyle =
            (style?.Bold == true ? OfficeFontStyle.Bold : OfficeFontStyle.Regular) |
            (style?.Italic == true ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        OfficeImageExportDiagnostic? diagnostic = fonts.CreateSubstitutionDiagnostic(
            text,
            family,
            fontStyle,
            source);
        if (diagnostic == null) return;
        string key = diagnostic.Code + "\n" + diagnostic.Message;
        if (seen.Add(key)) diagnostics.Add(diagnostic);
    }
}
