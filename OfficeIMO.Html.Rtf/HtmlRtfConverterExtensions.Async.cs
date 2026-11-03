using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

/// <content>
/// Provides asynchronous IO extension methods for converting between RTF and semantic HTML.
/// </content>
public static partial class HtmlRtfConverterExtensions {
    /// <summary>Saves an RTF document as semantic HTML at the specified path.</summary>
    public static async Task SaveAsHtmlAsync(this RtfDocument document, string path, RtfHtmlExportOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = (encoding ?? Encoding.UTF8).GetBytes(document.ToHtml(options));
        await OfficeFileCommit.WriteAllBytesAsync(path, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves an RTF document as semantic HTML to a writable stream.</summary>
    public static async Task SaveAsHtmlAsync(this RtfDocument document, Stream stream, RtfHtmlExportOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        cancellationToken.ThrowIfCancellationRequested();
        await OfficeStreamWriter.WriteAllBytesAsync(stream, document.ToHtmlBytes(options, encoding), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves semantic HTML as an RTF file at the specified path.</summary>
    public static async Task SaveAsRtfAsync(this HtmlConversionDocument document, string path, HtmlRtfExportOptions? readOptions = null, RtfSaveOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        RtfDocument rtfDocument = document.ToRtfDocument(readOptions);
        await rtfDocument.SaveAsync(path, writeOptions, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves semantic HTML as RTF to a writable stream without closing or rewinding the stream.</summary>
    public static async Task SaveAsRtfAsync(this HtmlConversionDocument document, Stream stream, HtmlRtfExportOptions? readOptions = null, RtfSaveOptions? writeOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        RtfDocument rtfDocument = document.ToRtfDocument(readOptions);
        await rtfDocument.SaveAsync(stream, writeOptions, encoding, cancellationToken).ConfigureAwait(false);
    }
}
