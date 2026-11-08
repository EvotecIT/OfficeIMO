using OfficeIMO.Html;

namespace OfficeIMO.Word.Html;

public static partial class WordHtmlConverterExtensions {
    private static HtmlConversionDocument ApplyThemeDefaultsForWord(HtmlConversionDocument document, HtmlWordExportOptions options) {
        if (options.Theme is not { } theme) return document;
        if (!string.IsNullOrWhiteSpace(options.FontFamily)) {
            // An adapter-selected default face wins the common theme; authored CSS remains
            // above the generated default layer and keeps its own font choices.
            theme = theme.WithFonts(theme.Fonts.With(
                heading: theme.Fonts.Heading.With(family: options.FontFamily),
                body: theme.Fonts.Body.With(family: options.FontFamily)));
        }
        if (!theme.IsDark) return document.WithTheme(theme);
        // Native Word keeps a light page even for screen-oriented dark themes.
        // These remain defaults: authored HTML paint stays above this layer.
        string pageDefaults = ":where(html),:where(html) body{color:#" + WordThemeDefaults.PageText(theme).ToRgbHex() + ";background:#FFFFFF;color-scheme:light;}\n" +
            ":where(html) h1,:where(html) h2,:where(html) h3,:where(html) h4,:where(html) h5,:where(html) h6{color:#" + WordThemeDefaults.PageTextColor(theme.TextStyle.HeadingColor, theme).ToRgbHex() + ";}\n" +
            ":where(html) pre{background:#" + OfficeTheme.Modern.Surfaces.Surface.Resolve(theme).ToRgbHex() + ";}\n" +
            ":where(html) tbody>tr:nth-child(even)>td{background:#" + OfficeTheme.Modern.TableStyle.BandFill.Resolve(theme).ToRgbHex() + ";}\n";
        return document.WithTheme(theme, pageDefaults);
    }
}
