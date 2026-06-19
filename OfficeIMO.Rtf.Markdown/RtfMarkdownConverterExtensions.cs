using System;
using System.IO;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>
/// Semantic conversion helpers between OfficeIMO RTF and Markdown document models.
/// </summary>
public static partial class RtfMarkdownConverterExtensions {
    /// <summary>
    /// Converts an RTF document into a Markdown document model.
    /// </summary>
    public static MarkdownDoc ToMarkdownDocument(this RtfDocument document, RtfToMarkdownOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return RtfToMarkdownConverter.Convert(document, options ?? new RtfToMarkdownOptions());
    }

    /// <summary>
    /// Converts an RTF document into Markdown text.
    /// </summary>
    public static string ToMarkdown(this RtfDocument document, RtfToMarkdownOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var effectiveOptions = options ?? new RtfToMarkdownOptions();
        return document.ToMarkdownDocument(effectiveOptions).ToMarkdown(effectiveOptions.MarkdownWriteOptions);
    }

    /// <summary>
    /// Writes an RTF document as Markdown text.
    /// </summary>
    public static void SaveAsMarkdown(this RtfDocument document, string path, RtfToMarkdownOptions? options = null, Encoding? encoding = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path is required.", nameof(path));

        File.WriteAllText(path, document.ToMarkdown(options), encoding ?? new UTF8Encoding(false));
    }

    /// <summary>
    /// Converts a Markdown document model into an RTF document.
    /// </summary>
    public static RtfDocument ToRtfDocument(this MarkdownDoc markdown, MarkdownToRtfOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        return MarkdownToRtfConverter.Convert(markdown, options ?? new MarkdownToRtfOptions());
    }

    /// <summary>
    /// Parses Markdown text and converts it into an RTF document.
    /// </summary>
    public static RtfDocument ToRtfDocumentFromMarkdown(this string markdown, MarkdownToRtfOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));

        var effectiveOptions = options ?? new MarkdownToRtfOptions();
        var doc = MarkdownReader.Parse(markdown, effectiveOptions.ReaderOptions);
        return doc.ToRtfDocument(effectiveOptions);
    }

    /// <summary>
    /// Parses Markdown text, converts it into an RTF document, and renders RTF text.
    /// </summary>
    public static string ToRtfFromMarkdown(this string markdown, MarkdownToRtfOptions? options = null, RtfWriteOptions? writeOptions = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        return markdown.ToRtfDocumentFromMarkdown(options).ToRtf(writeOptions);
    }
}
