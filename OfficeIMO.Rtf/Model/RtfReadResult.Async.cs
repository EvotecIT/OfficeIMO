namespace OfficeIMO.Rtf;

/// <content>
/// Provides asynchronous lossless save APIs for read results.
/// </content>
public sealed partial class RtfReadResult {
    /// <summary>
    /// Serializes the original syntax tree to source-preserving bytes without semantic normalization.
    /// </summary>
    public Task<byte[]> ToBytesLosslessAsync(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(ToBytesLossless());
    }

    /// <summary>
    /// Saves the original RTF stream to a file without semantic normalization.
    /// </summary>
    public Task SaveLosslessAsync(string path, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return RtfBytePreservingEncoding.WriteAllTextAsync(path, ToRtfLossless(), cancellationToken);
    }

    /// <summary>
    /// Saves the original RTF stream to a stream without semantic normalization.
    /// </summary>
    public Task SaveLosslessAsync(Stream stream, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return RtfBytePreservingEncoding.WriteToAsync(stream, ToRtfLossless(), cancellationToken);
    }
}
