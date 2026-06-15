namespace OfficeIMO.Word.Rtf;

/// <content>
/// Provides asynchronous Word and RTF serialization input and output APIs.
/// </content>
public static partial class WordRtfConverterExtensions {
    /// <summary>Serializes a Word document to RTF text asynchronously.</summary>
    public static Task<string> ToRtfAsync(this WordDocument document, RtfWriteOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        return document.ToRtfDocument().ToRtfAsync(options, cancellationToken);
    }

    /// <summary>Serializes a Word document to encoded RTF bytes asynchronously.</summary>
    public static async Task<byte[]> ToRtfBytesAsync(this WordDocument document, RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        return await document.ToRtfDocument().ToBytesAsync(options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Serializes a Word document to an encoded RTF memory stream asynchronously.</summary>
    public static async Task<MemoryStream> ToRtfMemoryStreamAsync(this WordDocument document, RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        return await document.ToRtfDocument().ToMemoryStreamAsync(options, encoding, cancellationToken).ConfigureAwait(false);
    }

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

    /// <summary>Creates a Word document from RTF text asynchronously.</summary>
    public static async Task<WordDocument> LoadFromRtfAsync(this string rtf, RtfReadOptions? readOptions = null, CancellationToken cancellationToken = default) {
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));
        RtfReadResult result = await RtfDocument.ReadAsync(rtf, readOptions, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return result.Document.ToWordDocument();
    }

    /// <summary>Creates a Word document from source RTF bytes asynchronously using the core byte-preserving RTF reader.</summary>
    public static async Task<WordDocument> LoadFromRtfAsync(this byte[] rtfBytes, RtfReadOptions? readOptions = null, CancellationToken cancellationToken = default) {
        if (rtfBytes == null) throw new ArgumentNullException(nameof(rtfBytes));
        RtfReadResult result = await RtfDocument.LoadAsync(rtfBytes, readOptions, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return result.Document.ToWordDocument();
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
