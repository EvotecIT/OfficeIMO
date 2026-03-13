using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Extension helpers for converting HTML into OfficeIMO.Markdown content.
/// </summary>
public static class HtmlMarkdownConverterExtensions {
    /// <summary>
    /// Converts an HTML string into Markdown text.
    /// </summary>
    public static string ToMarkdown(this string html, HtmlToMarkdownOptions? options = null) {
        var converter = new HtmlToMarkdownConverter();
        return converter.Convert(html, options ?? new HtmlToMarkdownOptions());
    }

    /// <summary>
    /// Converts an HTML stream into Markdown text.
    /// </summary>
    public static string ToMarkdown(this Stream htmlStream, HtmlToMarkdownOptions? options = null) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
        return reader.ReadToEnd().ToMarkdown(options);
    }

    /// <summary>
    /// Converts an HTML string into a Markdown document model.
    /// </summary>
    public static MarkdownDoc LoadFromHtml(this string html, HtmlToMarkdownOptions? options = null) {
        var converter = new HtmlToMarkdownConverter();
        return converter.ConvertToDocument(html, options ?? new HtmlToMarkdownOptions());
    }

    /// <summary>
    /// Converts an HTML stream into a Markdown document model.
    /// </summary>
    public static MarkdownDoc LoadFromHtml(this Stream htmlStream, HtmlToMarkdownOptions? options = null) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
        return reader.ReadToEnd().LoadFromHtml(options);
    }
}
