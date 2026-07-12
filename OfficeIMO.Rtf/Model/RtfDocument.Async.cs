namespace OfficeIMO.Rtf;

/// <content>
/// Provides asynchronous RTF document read and save APIs.
/// </content>
public sealed partial class RtfDocument {
    /// <summary>Reads RTF from a string.</summary>
    public static Task<RtfReadResult> ReadAsync(string rtf, RtfReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (rtf == null) throw new ArgumentNullException(nameof(rtf));
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(Read(rtf, options, cancellationToken));
    }

    /// <summary>Loads RTF from a file.</summary>
    public static async Task<RtfReadResult> LoadAsync(string path, RtfReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize: 4096, useAsync: true);
        return await LoadAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Loads RTF from source bytes using the byte-preserving lossless representation.</summary>
    public static Task<RtfReadResult> LoadAsync(byte[] bytes, RtfReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        cancellationToken.ThrowIfCancellationRequested();
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        new RtfReadLimitGuard(readOptions, cancellationToken).CheckInputBytes(bytes.LongLength);
        return Task.FromResult(Read(RtfBytePreservingEncoding.GetString(bytes), readOptions, cancellationToken));
    }

    /// <summary>Loads RTF from a stream.</summary>
    public static async Task<RtfReadResult> LoadAsync(Stream stream, RtfReadOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        RtfReadOptions readOptions = options ?? RtfReadOptions.CreateOfficeIMOProfile();
        byte[] bytes = await RtfBytePreservingEncoding.ReadBytesToEndAsync(stream, readOptions.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        string rtf = DecodeInput(bytes, encoding);
        return Read(rtf, readOptions, cancellationToken);
    }

    /// <summary>Serializes the document to RTF.</summary>
    public Task<string> ToRtfAsync(RtfWriteOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(ToRtf(options));
    }

    /// <summary>Serializes the document to encoded RTF bytes.</summary>
    public async Task<byte[]> ToBytesAsync(RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        string rtf = await ToRtfAsync(options, cancellationToken).ConfigureAwait(false);
        return (encoding ?? CreateDefaultOutputEncoding()).GetBytes(rtf);
    }

    /// <summary>Serializes the document to an encoded RTF memory stream.</summary>
    public async Task<MemoryStream> ToMemoryStreamAsync(RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        byte[] bytes = await ToBytesAsync(options, encoding, cancellationToken).ConfigureAwait(false);
        return new MemoryStream(bytes, writable: false);
    }

    /// <summary>Saves the document to an RTF file.</summary>
    public async Task SaveAsync(string path, RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        byte[] bytes = ToBytes(options, encoding);
        await OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytesAsync(path, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves the document to an RTF stream without closing the stream.</summary>
    public async Task SaveAsync(Stream stream, RtfWriteOptions? options = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        byte[] bytes = ToBytes(options, encoding);
        await OfficeIMO.Core.Internal.OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
