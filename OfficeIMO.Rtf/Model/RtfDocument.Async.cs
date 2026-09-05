using OfficeIMO.Core.Internal;
namespace OfficeIMO.Rtf;

/// <content>
/// Provides asynchronous RTF document read and save APIs.
/// </content>
public sealed partial class RtfDocument {
    /// <summary>Asynchronously loads RTF from a file into the semantic document model.</summary>
    public static async Task<RtfDocument> LoadAsync(string path, RtfReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) =>
        (await LoadResultAsync(path, options, encoding, cancellationToken).ConfigureAwait(false)).Document;

    /// <summary>Asynchronously loads RTF from a file with lossless syntax and diagnostics.</summary>
    public static async Task<RtfReadResult> LoadResultAsync(string path, RtfReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 4096, useAsync: true);
        return await LoadResultAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously loads RTF from a stream into the semantic document model.</summary>
    public static async Task<RtfDocument> LoadAsync(Stream stream, RtfReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) =>
        (await LoadResultAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false)).Document;

    /// <summary>Asynchronously loads RTF from a stream with lossless syntax and diagnostics.</summary>
    public static async Task<RtfReadResult> LoadResultAsync(Stream stream, RtfReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        byte[] bytes = await RtfBytePreservingEncoding.ReadBytesToEndAsync(stream, readOptions.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        string rtf = DecodeInput(bytes, encoding);
        return ParseResult(rtf, readOptions, cancellationToken).AttachOriginalBytes(bytes);
    }

    /// <summary>Saves the document to an RTF file.</summary>
    public async Task SaveAsync(string path, RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        byte[] bytes = ToFileBytes(options, encoding);
        await OfficeFileCommit.WriteAllBytesAsync(path, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves the document to an RTF stream without closing the stream.</summary>
    public async Task SaveAsync(Stream stream, RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        byte[] bytes = ToBytes(options, encoding);
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
