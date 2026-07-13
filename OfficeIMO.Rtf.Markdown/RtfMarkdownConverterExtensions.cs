using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>
/// Semantic conversion helpers between OfficeIMO RTF and Markdown document models.
/// </summary>
public static partial class RtfMarkdownConverterExtensions {
    /// <summary>
    /// Converts an RTF document into a Markdown document model.
    /// </summary>
    public static MarkdownDoc ToMarkdownDocument(this RtfDocument document, RtfToMarkdownOptions? options = null) {
        return document.ToMarkdownDocumentResult(options).Value;
    }

    /// <summary>Converts an RTF document into Markdown together with per-operation fidelity diagnostics.</summary>
    public static RtfConversionResult<MarkdownDoc> ToMarkdownDocumentResult(this RtfDocument document, RtfToMarkdownOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var context = new RtfToMarkdownConversionContext(options ?? new RtfToMarkdownOptions());
        MarkdownDoc value = RtfToMarkdownConverter.Convert(document, context);
        return new RtfConversionResult<MarkdownDoc>(value, context.ConversionReport);
    }

    /// <summary>
    /// Converts an RTF document into Markdown text.
    /// </summary>
    public static string ToMarkdown(this RtfDocument document, RtfToMarkdownOptions? options = null) {
        return document.ToMarkdownResult(options).Value;
    }

    /// <summary>Converts an RTF document into Markdown text with per-operation fidelity diagnostics.</summary>
    public static RtfConversionResult<string> ToMarkdownResult(this RtfDocument document, RtfToMarkdownOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var effectiveOptions = options ?? new RtfToMarkdownOptions();
        RtfConversionResult<MarkdownDoc> converted = document.ToMarkdownDocumentResult(effectiveOptions);
        string value = converted.Value.ToMarkdown(effectiveOptions.MarkdownWriteOptions);
        return new RtfConversionResult<string>(value, converted.Report);
    }

    /// <summary>Converts an RTF document into encoded Markdown bytes.</summary>
    public static byte[] ToMarkdownBytes(
        this RtfDocument document,
        RtfToMarkdownOptions? options = null,
        Encoding? encoding = null) =>
        (encoding ?? new UTF8Encoding(false)).GetBytes(document.ToMarkdown(options));

    /// <summary>Converts an RTF document into a writable Markdown stream positioned at the beginning.</summary>
    public static MemoryStream ToMarkdownStream(
        this RtfDocument document,
        RtfToMarkdownOptions? options = null,
        Encoding? encoding = null) =>
        new MemoryStream(document.ToMarkdownBytes(options, encoding));

    /// <summary>
    /// Writes an RTF document as Markdown text.
    /// </summary>
    public static void SaveAsMarkdown(this RtfDocument document, string path, RtfToMarkdownOptions? options = null, Encoding? encoding = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path is required.", nameof(path));

        OfficeFileCommit.WriteAllBytes(path, document.ToMarkdownBytes(options, encoding));
    }

    /// <summary>Writes an RTF document as Markdown to a caller-owned stream.</summary>
    public static void SaveAsMarkdown(
        this RtfDocument document,
        Stream stream,
        RtfToMarkdownOptions? options = null,
        Encoding? encoding = null) =>
        OfficeStreamWriter.WriteAllBytes(stream, document.ToMarkdownBytes(options, encoding));

    /// <summary>Asynchronously writes an RTF document as Markdown to a path.</summary>
    public static Task SaveAsMarkdownAsync(
        this RtfDocument document,
        string path,
        RtfToMarkdownOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) =>
        OfficeFileCommit.WriteAllBytesAsync(path, document.ToMarkdownBytes(options, encoding), cancellationToken: cancellationToken);

    /// <summary>Asynchronously writes an RTF document as Markdown to a caller-owned stream.</summary>
    public static Task SaveAsMarkdownAsync(
        this RtfDocument document,
        Stream stream,
        RtfToMarkdownOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) =>
        OfficeStreamWriter.WriteAllBytesAsync(stream, document.ToMarkdownBytes(options, encoding), cancellationToken);

    /// <summary>
    /// Converts a Markdown document model into an RTF document.
    /// </summary>
    public static RtfDocument ToRtfDocument(this MarkdownDoc markdown, MarkdownToRtfOptions? options = null) {
        return markdown.ToRtfDocumentResult(options).Value;
    }

    /// <summary>Converts Markdown into an RTF document together with per-operation fidelity diagnostics.</summary>
    public static RtfConversionResult<RtfDocument> ToRtfDocumentResult(this MarkdownDoc markdown, MarkdownToRtfOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        var context = new MarkdownToRtfConversionContext(options ?? new MarkdownToRtfOptions());
        RtfDocument value = MarkdownToRtfConverter.Convert(markdown, context);
        return new RtfConversionResult<RtfDocument>(value, context.ConversionReport);
    }

    /// <summary>Converts a parsed Markdown document into RTF text.</summary>
    public static string ToRtf(this MarkdownDoc markdown, MarkdownToRtfOptions? options = null, RtfWriteOptions? writeOptions = null) {
        return markdown.ToRtfResult(options, writeOptions).Value;
    }

    /// <summary>Converts a parsed Markdown document into RTF text with per-operation fidelity diagnostics.</summary>
    public static RtfConversionResult<string> ToRtfResult(
        this MarkdownDoc markdown,
        MarkdownToRtfOptions? options = null,
        RtfWriteOptions? writeOptions = null) {
        RtfConversionResult<RtfDocument> converted = markdown.ToRtfDocumentResult(options);
        return new RtfConversionResult<string>(converted.Value.ToRtf(writeOptions), converted.Report);
    }

    /// <summary>Converts a parsed Markdown document into encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(
        this MarkdownDoc markdown,
        MarkdownToRtfOptions? options = null,
        RtfWriteOptions? writeOptions = null,
        Encoding? encoding = null) =>
        (encoding ?? new UTF8Encoding(false)).GetBytes(markdown.ToRtf(options, writeOptions));

    /// <summary>Converts a parsed Markdown document into a writable RTF stream positioned at the beginning.</summary>
    public static MemoryStream ToRtfStream(
        this MarkdownDoc markdown,
        MarkdownToRtfOptions? options = null,
        RtfWriteOptions? writeOptions = null,
        Encoding? encoding = null) =>
        new MemoryStream(markdown.ToRtfBytes(options, writeOptions, encoding));

    /// <summary>Saves a parsed Markdown document as RTF to a path.</summary>
    public static void SaveAsRtf(
        this MarkdownDoc markdown,
        string path,
        MarkdownToRtfOptions? options = null,
        RtfWriteOptions? writeOptions = null,
        Encoding? encoding = null) =>
        OfficeFileCommit.WriteAllBytes(path, markdown.ToRtfBytes(options, writeOptions, encoding));

    /// <summary>Saves a parsed Markdown document as RTF to a caller-owned stream.</summary>
    public static void SaveAsRtf(
        this MarkdownDoc markdown,
        Stream stream,
        MarkdownToRtfOptions? options = null,
        RtfWriteOptions? writeOptions = null,
        Encoding? encoding = null) =>
        OfficeStreamWriter.WriteAllBytes(stream, markdown.ToRtfBytes(options, writeOptions, encoding));

    /// <summary>Asynchronously saves a parsed Markdown document as RTF to a path.</summary>
    public static Task SaveAsRtfAsync(
        this MarkdownDoc markdown,
        string path,
        MarkdownToRtfOptions? options = null,
        RtfWriteOptions? writeOptions = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) =>
        OfficeFileCommit.WriteAllBytesAsync(path, markdown.ToRtfBytes(options, writeOptions, encoding), cancellationToken: cancellationToken);

    /// <summary>Asynchronously saves a parsed Markdown document as RTF to a caller-owned stream.</summary>
    public static Task SaveAsRtfAsync(
        this MarkdownDoc markdown,
        Stream stream,
        MarkdownToRtfOptions? options = null,
        RtfWriteOptions? writeOptions = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) =>
        OfficeStreamWriter.WriteAllBytesAsync(stream, markdown.ToRtfBytes(options, writeOptions, encoding), cancellationToken);
}
