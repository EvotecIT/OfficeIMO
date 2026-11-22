using System.Globalization;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool AreNativeOpenTypeFeaturesEnabled(WordDocument? document) {
        W.Settings? settings = document?._wordprocessingDocument?.MainDocumentPart?.DocumentSettingsPart?.Settings;
        string? value = settings?.GetFirstChild<W.Compatibility>()?.Elements<W.CompatibilitySetting>()
            .LastOrDefault(setting => setting.Name?.InnerText == "enableOpenTypeFeatures" &&
                setting.Uri?.Value == "http://schemas.microsoft.com/office/word")?.Val?.Value;
        if (value == "0" || value == "false" || value == "off") return false;
        if (value == "1" || value == "true" || value == "on") return true;
        return document != null && UsesModernNativeWordLayout(document);
    }

    // Word 2013 and later drop the space before a paragraph that starts a page through
    // page-break-before. Earlier compatibility modes keep it, reduced by the space after
    // the paragraph that ended the previous page.
    private static void RenderNativePageBreakBefore(INativePdfFlow pdf, WordParagraph paragraph) {
        if (!UsesModernNativeWordLayout(paragraph._document) && pdf is NativeSpacingCollapseFlow collapseFlow) {
            collapseFlow.PageBreakPreservingSpacingBefore();
            return;
        }

        pdf.PageBreak();
    }

    /// <summary>
    /// Word 2010 and earlier compatibility modes with "Suppress extra line spacing at top of page
    /// (WordPerfect)" set every automatic line to the font size plus two points of leading, or
    /// the font size alone for a fixed-pitch font, instead of the font's own line metrics.
    /// Word 2013 mode ignores the option.
    /// </summary>
    private static bool UsesNativeWordPerfectLineHeight(WordDocument? document) =>
        document != null && !UsesModernNativeWordLayout(document) &&
        document._wordprocessingDocument?.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<W.Compatibility>()?.GetFirstChild<W.SuppressTopSpacingWordPerfect>() is { } option &&
        option.Val?.Value != false;

    private const double NativeWordPerfectLeading = 2D;

    private static double ResolveNativeWordPerfectLineHeight(WordParagraph paragraph, NativeDocumentDefaults defaults,
        NativeParagraphStyleDefaults styleDefaults, NativeTableRunStyleDefaults tableRunStyleDefaults, NativeFontMap? fontMap) {
        double fontSize = ResolveNativeParagraphEffectiveFontSize(paragraph, defaults, styleDefaults, tableRunStyleDefaults);
        if (fontSize <= 0D) return NativeDefaultParagraphLineHeight;
        string? family = EnumerateNativeStyleFontFamilies(default, styleDefaults, tableRunStyleDefaults, defaults, true).FirstOrDefault();
        foreach (WordParagraph run in GetNativeRuns(paragraph)) {
            if (!IsNativeParagraphMetricBearingRun(run, paragraph, defaults.NoteContentSection != null)) continue;
            family = EnumerateNativeParagraphOwnFontFamilies(run).FirstOrDefault() ?? family;
            break;
        }
        return (fontSize + (IsNativeFixedPitchFamily(family, fontMap) ? 0D : NativeWordPerfectLeading)) / fontSize;
    }

    private static bool IsNativeFixedPitchFamily(string? family, NativeFontMap? fontMap) {
        if (string.IsNullOrWhiteSpace(family)) return false;
        if (fontMap?.TryGetFontSlot(family, out OfficeIMO.Pdf.PdfStandardFont slot) == true &&
            OfficeIMO.Pdf.PdfStandardFontMapper.GetFontFamily(slot) == OfficeIMO.Pdf.PdfStandardFont.Courier) return true;
        string name = family!.ToLowerInvariant();
        return name.Contains("courier") || name.Contains("consolas") || name.Contains("lucida console") ||
            name.Contains(" mono") || name.EndsWith("mono", StringComparison.Ordinal);
    }

    /// <summary>Resolves the effective Word layout mode while retaining stored compatibility flags.</summary>
    private static bool UsesModernNativeWordLayout(WordDocument document) {
        if (document.SourceFormat == WordFileFormat.Doc) return false;
        W.Settings? settings = document._wordprocessingDocument?.MainDocumentPart?.DocumentSettingsPart?.Settings;
        string? modeText = settings?.GetFirstChild<W.Compatibility>()?.Elements<W.CompatibilitySetting>()
            .LastOrDefault(setting => setting.Name?.Value == W.CompatSettingNameValues.CompatibilityMode)?.Val?.Value;
        return int.TryParse(modeText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int mode) && mode >= 15;
    }
}
