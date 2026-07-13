using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Rtf;

/// <content>
/// Provides asynchronous RTF document read and save APIs.
/// </content>
public sealed partial class RtfDocument {
    /// <summary>Loads RTF from a file.</summary>
    public static async Task<RtfReadResult> LoadAsync(string path, RtfReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 4096, useAsync: true);
        return await LoadAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Loads RTF from a stream.</summary>
    public static async Task<RtfReadResult> LoadAsync(Stream stream, RtfReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        byte[] bytes = await RtfBytePreservingEncoding.ReadBytesToEndAsync(stream, readOptions.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        string rtf = DecodeInput(bytes, encoding);
        return Read(rtf, readOptions, cancellationToken);
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
