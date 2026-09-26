using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static string OmitUnavailablePrivateUseGlyphs(
        HtmlRenderText visual,
        RegisteredWebFonts webFonts,
        PdfCore.PdfConversionReport conversionReport,
        OfficeFontStyle requestedStyle,
        CancellationToken cancellationToken) {
        string text = visual.Text;
        StringBuilder? retained = null;
        for (int index = 0; index < text.Length; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            int scalarLength = ScalarLength(text, index);
            if (CharUnicodeInfo.GetUnicodeCategory(text, index) == UnicodeCategory.PrivateUse) {
                string glyph = text.Substring(index, scalarLength);
                if (!CanPaintPrivateUseGlyph(visual, glyph, requestedStyle, webFonts)) {
                    retained ??= new StringBuilder(text.Substring(0, index));
                    int scalar = scalarLength == 2 ? char.ConvertToUtf32(text, index) : text[index];
                    string codePoint = "U+" + scalar.ToString("X4", CultureInfo.InvariantCulture);
                    string family = visual.Font.FamilyName ?? string.Empty;
                    if (webFonts.ReportedPrivateUseOmissions.Add(family + "\0" + codePoint)) {
                        conversionReport.Add(new PdfCore.PdfConversionWarning(
                            "OfficeIMO.Html.Pdf",
                            HtmlPdfDiagnosticCodes.UnavailablePrivateUseGlyphOmitted,
                            visual.Source ?? "html-text",
                            "A private-use glyph was omitted because no usable PDF font or outline path was available.",
                            PdfCore.PdfConversionWarningSeverity.Warning,
                            OfficeConversionLossKind.Omission,
                            details: new Dictionary<string, string> {
                                ["CodePoint"] = codePoint,
                                ["FontFamily"] = family
                            }));
                    }
                    index += scalarLength - 1;
                    continue;
                }
            }
            if (retained != null) retained.Append(text, index, scalarLength);
            index += scalarLength - 1;
        }
        return retained?.ToString() ?? text;
    }

    private static string? FilterLogicalPrivateUseGlyphs(
        string logicalText,
        IEnumerable<HtmlRenderVisual> visuals,
        RegisteredWebFonts webFonts,
        CancellationToken cancellationToken) {
        if (logicalText.Length == 0) return logicalText;
        var paintability = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (HtmlRenderText child in EnumerateVisuals(visuals).OfType<HtmlRenderText>()) {
            OfficeFontStyle style = (child.Font.IsBold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
                | (child.Font.IsItalic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
            for (int index = 0; index < child.Text.Length; index++) {
                if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                int length = ScalarLength(child.Text, index);
                if (CharUnicodeInfo.GetUnicodeCategory(child.Text, index) == UnicodeCategory.PrivateUse) {
                    string glyph = child.Text.Substring(index, length);
                    bool canPaint = CanPaintPrivateUseGlyph(child, glyph, style, webFonts);
                    int state = canPaint ? 1 : 2;
                    paintability[glyph] = state | (paintability.TryGetValue(glyph, out int previous) ? previous : 0);
                }
                index += length - 1;
            }
        }
        // A mixed-font occurrence cannot be mapped back to the logical string by code point.
        // Let the painted child runs own extraction rather than claim an omitted glyph in ActualText.
        if (paintability.Values.Any(static value => value == 3)) return null;
        if (!paintability.Values.Any(static value => value == 2)) return logicalText;
        StringBuilder? retained = null;
        for (int index = 0; index < logicalText.Length; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            int length = ScalarLength(logicalText, index);
            if (CharUnicodeInfo.GetUnicodeCategory(logicalText, index) == UnicodeCategory.PrivateUse
                && paintability.TryGetValue(logicalText.Substring(index, length), out int state)
                && state == 2) {
                retained ??= new StringBuilder(logicalText.Substring(0, index));
            } else if (retained != null) {
                retained.Append(logicalText, index, length);
            }
            index += length - 1;
        }
        return retained?.ToString() ?? logicalText;
    }

    private static bool ContainsPdfRenderableVisual(
        HtmlRenderVisual visual,
        RegisteredWebFonts webFonts,
        double surfaceWidth,
        double surfaceHeight,
        ClipBounds? activeClip,
        CancellationToken cancellationToken) {
        if (!ContainsRenderableVisual(visual, surfaceWidth, surfaceHeight, activeClip)) return false;
        if (visual is HtmlRenderText text) {
            OfficeFontStyle style = (text.Font.IsBold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
                | (text.Font.IsItalic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
            for (int index = 0; index < text.Text.Length; index++) {
                if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                int length = ScalarLength(text.Text, index);
                if (CharUnicodeInfo.GetUnicodeCategory(text.Text, index) != UnicodeCategory.PrivateUse
                    || CanPaintPrivateUseGlyph(text, text.Text.Substring(index, length), style, webFonts)) return true;
                index += length - 1;
            }
            return false;
        }
        IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderLayoutRegion region ? region.Visuals
            : visual is HtmlRenderSemanticGroup semantic ? semantic.Visuals
            : visual is HtmlRenderLogicalTextGroup logical ? logical.Visuals
            : visual is HtmlRenderClipGroup clip ? clip.Visuals
            : visual is HtmlRenderPathClipGroup pathClip ? pathClip.Visuals
            : visual is HtmlRenderEffectGroup effect ? effect.Visuals
            : null;
        return children == null || children.Any(child => ContainsPdfRenderableVisual(
            child, webFonts, surfaceWidth, surfaceHeight, activeClip, cancellationToken));
    }

    private static bool CanPaintPrivateUseGlyph(
        HtmlRenderText visual,
        string glyph,
        OfficeFontStyle style,
        RegisteredWebFonts webFonts) {
        string key = visual.Font.FamilyName + "\0" + ((int)style).ToString(CultureInfo.InvariantCulture)
            + "\0" + (visual.FeatureSettings.IsDefault ? "default" : "features") + "\0" + glyph;
        if (webFonts.PrivateUsePaintability.TryGetValue(key, out bool cached)) return cached;

        bool paintable = false;
        if (webFonts.Faces.TryResolveFaceForText(glyph, visual.Font.FamilyName, style, out OfficeFontFace? face)
            && face != null
            && (!face.CanEmbedAsStaticPdfFont || !visual.FeatureSettings.IsDefault
                || ContainsColorGlyph(face.Program, glyph))) {
            paintable = true;
        } else {
            OfficeFontFallbackRun? fallback = webFonts.Faces.PlanFallbackRuns(
                glyph, visual.Font.FamilyName, style).FirstOrDefault();
            string family = fallback?.FamilyName ?? visual.Font.FamilyName;
            if (webFonts.AllowInstalledFontFaces
                && !webFonts.Slots.ContainsKey(family)
                && EnumerateBoundedSystemFamilies(family).Any(candidate =>
                    NamedFontCoversTextWithStyleFallback(webFonts.Options, candidate, glyph,
                        visual.Font.IsBold, visual.Font.IsItalic))) {
                webFonts.PrivateUsePaintability[key] = true;
                return true;
            }
            var run = new PdfCore.PdfTextRun(
                glyph,
                bold: visual.Font.IsBold,
                italic: visual.Font.IsItalic,
                fontSize: visual.Font.Size * PointsPerCssPixel,
                font: MapFont(family, glyph, style, webFonts),
                fontFamily: family).WithFeatureSettings(visual.FeatureSettings);
            try {
                paintable = PdfCore.PdfWriter.MeasurePositionedText(run, webFonts.Options).HasValue;
            } catch (NotSupportedException) {
                paintable = false;
            } catch (ArgumentException) {
                paintable = false;
            }
        }
        webFonts.PrivateUsePaintability[key] = paintable;
        return paintable;
    }

    private static int ScalarLength(string text, int index) =>
        char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])
            ? 2 : 1;
}
