namespace OfficeIMO.Word.Rtf;

/// <content>
/// Provides asynchronous Word and RTF serialization input and output APIs.
/// </content>
public static partial class WordRtfConverterExtensions {
    /// <summary>Saves a Word document as an RTF file asynchronously.</summary>
    public static async Task SaveAsRtfAsync(this WordDocument document, string path, RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (path == null) throw new ArgumentNullException(nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        await document.ToRtfDocument().SaveAsync(path, options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves a Word document as RTF to a stream asynchronously without closing or rewinding the stream.</summary>
    public static async Task SaveAsRtfAsync(this WordDocument document, Stream stream, RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        await document.ToRtfDocument().SaveAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
    }

}
