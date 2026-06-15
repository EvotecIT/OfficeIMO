namespace OfficeIMO.Html;

/// <content>
/// Provides file-loading extension methods for converting semantic HTML into RTF documents.
/// </content>
public static partial class HtmlRtfConverterExtensions {
    /// <summary>Loads semantic HTML from a file into an RTF document model.</summary>
    public static RtfDocument ToRtfDocumentFromHtmlFile(string path, HtmlToRtfOptions? options = null, Encoding? encoding = null) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return stream.ToRtfDocument(options, encoding);
    }

    /// <summary>Loads semantic HTML from a file into an RTF document model asynchronously.</summary>
    public static async Task<RtfDocument> ToRtfDocumentFromHtmlFileAsync(string path, HtmlToRtfOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
        return await stream.ToRtfDocumentAsync(options, encoding, cancellationToken).ConfigureAwait(false);
    }
}
