using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal static partial class WordDocumentImageRenderer {
    private static T? ResolveImageRunProperty<T>(WordParagraph paragraph) where T : OpenXmlElement {
        T? direct = paragraph.ScopedRunProperties?.GetFirstChild<T>();
        if (direct != null) return direct;
        foreach (StyleRunProperties properties in EnumerateRunStyleProperties(paragraph)) {
            T? inherited = properties.GetFirstChild<T>();
            if (inherited != null) return inherited;
        }
        return ResolveDocumentDefaultRunProperties(paragraph)?.GetFirstChild<T>();
    }

    private static double ResolveImageFontSize(WordParagraph paragraph) =>
        double.TryParse(ResolveImageRunProperty<FontSize>(paragraph)?.Val?.Value,
            System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture,
            out double halfPoints) && halfPoints > 0D ? halfPoints / 2D : 11D;

    private static string ResolveImageFontFamily(WordParagraph paragraph) {
        RunFonts? fonts = ResolveImageRunProperty<RunFonts>(paragraph);
        string? family = fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
        if (!string.IsNullOrWhiteSpace(family)) return family!;
        string? theme = fonts?.AsciiTheme?.InnerText ?? fonts?.HighAnsiTheme?.InnerText;
        var scheme = paragraph._document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.FontScheme;
        family = theme?.StartsWith("major", StringComparison.Ordinal) == true
            ? scheme?.MajorFont?.LatinFont?.Typeface?.Value
            : scheme?.MinorFont?.LatinFont?.Typeface?.Value;
        return string.IsNullOrWhiteSpace(family) ? "Calibri" : family!;
    }

    private static bool ResolveImageBold(WordParagraph paragraph) =>
        ResolveImageExportOnOff(paragraph.ScopedRunProperties?.Bold, paragraph);

    private static bool ResolveImageItalic(WordParagraph paragraph) =>
        ResolveImageExportOnOff(paragraph.ScopedRunProperties?.Italic, paragraph);
}
