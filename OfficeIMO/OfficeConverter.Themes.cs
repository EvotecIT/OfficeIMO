using OfficeIMO.Markdown;

namespace OfficeIMO;

public static partial class OfficeConverter {
    /// <summary>Rejects appearance requests that the selected owner cannot carry into its artifact.</summary>
    internal static void ValidateThemeRoute(OfficeConversionOptions settings, string routeId) {
        if (settings.Theme == null) return;
        if (routeId.StartsWith("native-word-", StringComparison.OrdinalIgnoreCase) || routeId is
            "word-xml-docx" or "doc-pdf" or "docx-pdf" or "docx-html" or
            "markdown-docx" or "markdown-pdf" or "markdown-html" or "html-docx" or "html-pdf" or
            "docx-rtf" or "rtf-docx" or "rtf-pdf" or "rtf-html") return;
        throw new NotSupportedException("The selected conversion route '" + routeId + "' does not support a common Theme. Use the owning model's explicit styling API or omit Theme.");
    }

    private static MarkdownHtmlExportOptions ConfigureThemedMarkdownHtml(OfficeConversionOptions settings) {
        MarkdownHtmlExportOptions options = settings.MarkdownHtml ?? new MarkdownHtmlExportOptions {
            Kind = settings.Theme == null ? MarkdownHtmlKind.Fragment : MarkdownHtmlKind.Document
        };
        if (settings.Theme == null) return options;
        options.Theme ??= MarkdownVisualTheme.FromOfficeTheme(settings.Theme);
        if (options.CssDelivery == MarkdownCssDelivery.None || options.CssDelivery == MarkdownCssDelivery.ExternalFile ||
            options.Kind == MarkdownHtmlKind.Fragment && (!options.IncludeCssInFragment || options.CssDelivery != MarkdownCssDelivery.Inline))
            throw new NotSupportedException("A themed Markdown HTML conversion requires a document carrying its CSS, or an inline fragment with IncludeCssInFragment=true. Use the Markdown model API for host-owned or separately published stylesheets.");
        return options;
    }
}
