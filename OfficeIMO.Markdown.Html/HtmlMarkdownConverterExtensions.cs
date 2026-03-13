using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Extension helpers for converting HTML into OfficeIMO.Markdown content.
/// </summary>
public static class HtmlMarkdownConverterExtensions {
    /// <summary>
    /// Converts an HTML fragment or document string into Markdown text.
    /// </summary>
    /// <param name="html">HTML fragment or document to convert.</param>
    /// <param name="options">Optional conversion options. Default options are used when omitted.</param>
    /// <returns>The rendered Markdown text.</returns>
    public static string ToMarkdown(this string html, HtmlToMarkdownOptions? options = null) {
        var converter = new HtmlToMarkdownConverter();
        return converter.Convert(html, options);
    }

    /// <summary>
    /// Converts an HTML stream into Markdown text.
    /// </summary>
    /// <param name="htmlStream">Readable stream containing HTML markup.</param>
    /// <param name="options">Optional conversion options. Default options are used when omitted.</param>
    /// <returns>The rendered Markdown text.</returns>
    public static string ToMarkdown(this Stream htmlStream, HtmlToMarkdownOptions? options = null) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
        return reader.ReadToEnd().ToMarkdown(options);
    }

    /// <summary>
    /// Converts an HTML fragment or document string into a Markdown document model.
    /// </summary>
    /// <param name="html">HTML fragment or document to convert.</param>
    /// <param name="options">Optional conversion options. Default options are used when omitted.</param>
    /// <returns>A structural <see cref="MarkdownDoc"/> representing the converted Markdown.</returns>
    public static MarkdownDoc LoadFromHtml(this string html, HtmlToMarkdownOptions? options = null) {
        var converter = new HtmlToMarkdownConverter();
        return converter.ConvertToDocument(html, options);
    }

    /// <summary>
    /// Converts an HTML stream into a Markdown document model.
    /// </summary>
    /// <param name="htmlStream">Readable stream containing HTML markup.</param>
    /// <param name="options">Optional conversion options. Default options are used when omitted.</param>
    /// <returns>A structural <see cref="MarkdownDoc"/> representing the converted Markdown.</returns>
    public static MarkdownDoc LoadFromHtml(this Stream htmlStream, HtmlToMarkdownOptions? options = null) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        using var reader = new StreamReader(htmlStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
        return reader.ReadToEnd().LoadFromHtml(options);
    }
}
