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

    /// <summary>Creates a Word document from an RTF stream asynchronously, reading from the stream's current position.</summary>
    public static async Task<WordDocument> LoadFromRtfAsync(this Stream rtfStream, RtfReadOptions? readOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (rtfStream == null) throw new ArgumentNullException(nameof(rtfStream));
        RtfReadResult result = await RtfDocument.LoadAsync(rtfStream, readOptions, encoding, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return result.Document.ToWordDocument();
    }

    /// <summary>Loads an RTF file asynchronously and converts it to a Word document.</summary>
    public static async Task<WordDocument> LoadFromRtfFileAsync(string path, RtfReadOptions? readOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        RtfReadResult result = await RtfDocument.LoadAsync(path, readOptions, encoding, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return result.Document.ToWordDocument();
    }
}
