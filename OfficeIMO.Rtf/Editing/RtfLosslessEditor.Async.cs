namespace OfficeIMO.Rtf;

/// <content>
/// Provides asynchronous lossless editor output APIs.
/// </content>
public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Serializes the edited syntax tree to source-preserving bytes without semantic normalization.
    /// </summary>
    public Task<byte[]> ToBytesLosslessAsync(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(ToBytesLossless());
    }

    /// <summary>
    /// Saves the edited syntax tree to a file without semantic normalization.
    /// </summary>
    public Task SaveLosslessAsync(string path, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return RtfBytePreservingEncoding.WriteAllTextAsync(path, ToRtf(), cancellationToken);
    }

    /// <summary>
    /// Saves the edited syntax tree to a stream without semantic normalization.
    /// </summary>
    public Task SaveLosslessAsync(Stream stream, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return RtfBytePreservingEncoding.WriteToAsync(stream, ToRtf(), cancellationToken);
    }

    /// <summary>
    /// Reads the edited syntax tree into a fresh semantic model.
    /// </summary>
    public Task<RtfReadResult> ToReadResultAsync(RtfReadOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return RtfDocument.ReadAsync(ToRtf(), options, cancellationToken);
    }
}
