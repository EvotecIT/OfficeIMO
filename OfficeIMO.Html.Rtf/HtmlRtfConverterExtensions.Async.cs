using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

/// <content>
/// Provides asynchronous IO extension methods for converting between RTF and semantic HTML.
/// </content>
public static partial class HtmlRtfConverterExtensions {
    /// <summary>Saves an RTF document as semantic HTML at the specified path.</summary>
    public static async Task SaveAsHtmlAsync(this RtfDocument document, string path, RtfToHtmlOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);

        ArgumentNullException.ThrowIfNull(path);

        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = (encoding ?? Encoding.UTF8).GetBytes(document.ToHtml(options));
        await OfficeFileCommit.WriteAllBytesAsync(path, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves an RTF document as semantic HTML to a writable stream.</summary>
    public static async Task SaveAsHtmlAsync(this RtfDocument document, Stream stream, RtfToHtmlOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);

        ArgumentNullException.ThrowIfNull(stream);

        cancellationToken.ThrowIfCancellationRequested();
        await OfficeStreamWriter.WriteAllBytesAsync(stream, document.ToHtmlBytes(options, encoding), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves semantic HTML as an RTF file at the specified path.</summary>
    public static async Task SaveAsRtfAsync(this HtmlConversionDocument document, string path, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(path);

        RtfDocument rtfDocument = document.ToRtfDocument(readOptions);
        await rtfDocument.SaveAsync(path, writeOptions, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves semantic HTML as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static async Task SaveAsRtfAsync(this HtmlConversionDocument document, Stream stream, HtmlToRtfOptions? readOptions = null, RtfWriteOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);

        RtfDocument rtfDocument = document.ToRtfDocument(readOptions);
        await rtfDocument.SaveAsync(stream, writeOptions, encoding, cancellationToken).ConfigureAwait(false);
    }
}
