using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static class VisioImageExportFontDiagnostics {
    internal static void Append(
        VisioPage page,
        OfficeFontFaceCollection fonts,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        string source,
        bool renderText,
        bool renderConnectorLabels,
        VisioRenderLayerVisibility layerVisibility,
        System.Threading.CancellationToken cancellationToken = default) {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var textStyles = new VisioNativeTextStyleResolver(page.OwnerDocument, cancellationToken, diagnostics, source);
        if (renderText) {
            foreach (VisioShape shape in page.Shapes) {
                AppendShape(page, shape, fonts, diagnostics, seen, source, cancellationToken, textStyles, layerVisibility);
            }
        }
        if (renderConnectorLabels) {
            foreach (VisioConnector connector in page.Connectors) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!layerVisibility.IsVisible(connector)) continue;
                if (!AppendRuns(VisioRichTextProjection.Create(page, connector, 72D, cancellationToken, textStyles), fonts, diagnostics, seen, source))
                    AppendText(connector.Label, connector.TextStyle, fonts, diagnostics, seen, source);
            }
        }
    }

    private static void AppendShape(
        VisioPage page,
        VisioShape shape,
        OfficeFontFaceCollection fonts,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        HashSet<string> seen,
        string source,
        System.Threading.CancellationToken cancellationToken, VisioNativeTextStyleResolver textStyles, VisioRenderLayerVisibility layerVisibility) {
        cancellationToken.ThrowIfCancellationRequested();
        if (layerVisibility.IsVisible(shape) && !AppendRuns(VisioRichTextProjection.Create(page, shape, 72D, cancellationToken, textStyles), fonts, diagnostics, seen, source))
            AppendText(shape.Text, shape.TextStyle, fonts, diagnostics, seen, source);
        foreach (VisioShape child in shape.Children) {
            AppendShape(page, child, fonts, diagnostics, seen, source, cancellationToken, textStyles, layerVisibility);
        }
    }

    private static bool AppendRuns(VisioRichTextProjection? projection, OfficeFontFaceCollection fonts,
        ICollection<OfficeImageExportDiagnostic> diagnostics, HashSet<string> seen, string source) {
        if (projection == null) return false;
        foreach (OfficeRichTextRun run in projection.Runs) {
            OfficeImageExportDiagnostic? diagnostic = fonts.CreateSubstitutionDiagnostic(run.Text, run.FontFamily,
                (run.Bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular) |
                (run.Italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular), source);
            if (diagnostic != null && seen.Add(diagnostic.Code + "\n" + diagnostic.Message)) diagnostics.Add(diagnostic);
        }
        return true;
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
