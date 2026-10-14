using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private IOfficeFontProgram? ResolveTextFont(string? text, string? fontFamily, OfficeFontStyle style, double size) =>
        ResolveTextFont(text, fontFamily, style, size, out _);

    private IOfficeFontProgram? ResolveTextFont(string? text, string? fontFamily, OfficeFontStyle style, double size, out OfficeFontStyle resolvedStyle) {
        resolvedStyle = OfficeFontStyle.Regular;
        if (_fonts != null) {
            IOfficeFontProgram? scoped = _fonts.ResolveForText(text ?? string.Empty, fontFamily, RequestedTextFace(style), size / FontMetricScale, out resolvedStyle);
            if (scoped != null) {
                return ResolveMetricScale(scoped);
            }
        }

        if (_scopedFontResolutionOnly) return null;

        if (!string.IsNullOrWhiteSpace(fontFamily)) {
            OfficeTrueTypeFont? installed = OfficeTrueTypeFont.TryLoadFontFamilyForText(fontFamily, RequestedTextFace(style), text, out resolvedStyle);
            installed = installed?.ForInstalledOpticalSize(size / FontMetricScale, RequestedTextFace(style).Weight >= 600);
            if (installed?.HasSelectedBoldWeight == true) resolvedStyle |= OfficeFontStyle.Bold;
            if (installed != null) return ResolveMetricScale(installed);
        }

        // Measurement, outline bounds and painting must reject the same incomplete
        // default face; otherwise an empty .notdef outline silently loses text.
        resolvedStyle = OfficeFontStyle.Regular;
        if (_font != null && (string.IsNullOrEmpty(text) || _font.HasGlyphs(text!))) return ResolveMetricScale(_font);
        if (OfficeManagedTextShaper.RequiresComplexLayout(text)) ReportTextShapingFallback(incomplete: true);
        return null;
    }

    // Every candidate can select a different face, including its appended ellipsis.
    // Resolve coverage and metrics again before choosing the final painted string.
    private string FitRasterText(
        string text, double size, double availableWidth, string? fontFamily, OfficeFontStyle style,
        OfficeTextFeatureSettings? featureSettings, OfficeTextDirection textDirection) {
        if (MeasurePositionedText(text, size, fontFamily, style, featureSettings, textDirection) <= availableWidth) return text;
        string prefix = text;
        while (prefix.Length > 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            _textInkLayoutWork?.Invoke(prefix.Length);
            prefix = OfficeTextElements.RemoveLast(prefix);
            string candidate = prefix + "...";
            if (MeasurePositionedText(candidate, size, fontFamily, style, featureSettings, textDirection) <= availableWidth) return candidate;
        }
        return string.Empty;
    }

}
