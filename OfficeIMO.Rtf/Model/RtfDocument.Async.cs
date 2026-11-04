using OfficeIMO.Core.Internal;
namespace OfficeIMO.Rtf;

/// <content>
/// Provides asynchronous RTF document read and save APIs.
/// </content>
public sealed partial class RtfDocument {
    /// <summary>Asynchronously loads RTF from a file into the semantic document model.</summary>
    public static async Task<RtfDocument> LoadAsync(string path, RtfLoadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) =>
        (await LoadResultAsync(path, options, encoding, cancellationToken).ConfigureAwait(false)).Document;

    /// <summary>Asynchronously loads RTF from a file with lossless syntax and diagnostics.</summary>
    public static async Task<RtfReadResult> LoadResultAsync(string path, RtfLoadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        string fullPath = Path.GetFullPath(path);
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 4096, useAsync: true);
        RtfReadResult result = await LoadResultAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
        result.Document.FilePath = fullPath;
        return result;
    }

    /// <summary>Asynchronously loads RTF from a stream into the semantic document model.</summary>
    public static async Task<RtfDocument> LoadAsync(Stream stream, RtfLoadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) =>
        (await LoadResultAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false)).Document;

    /// <summary>Asynchronously loads RTF from a stream with lossless syntax and diagnostics.</summary>
    public static async Task<RtfReadResult> LoadResultAsync(Stream stream, RtfLoadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);
        RtfLoadOptions readOptions = options ?? new RtfLoadOptions();
        byte[] bytes = await RtfBytePreservingEncoding.ReadBytesToEndAsync(stream, readOptions.Limits.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        string rtf = DecodeInput(bytes, encoding);
        return ParseResult(rtf, readOptions, cancellationToken).AttachOriginalBytes(bytes);
    }

    /// <summary>Saves the document to an RTF file.</summary>
    public async Task SaveAsync(string path, RtfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        byte[] bytes = ToFileBytes(options, encoding);
        string fullPath = Path.GetFullPath(path);
        await OfficeFileCommit.WriteAllBytesAsync(fullPath, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
        FilePath = fullPath;
    }

    /// <summary>Saves the document to an RTF stream without closing the stream.</summary>
    public async Task SaveAsync(Stream stream, RtfSaveOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        byte[] bytes = ToBytes(options, encoding);
        await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
