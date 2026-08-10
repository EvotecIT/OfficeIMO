namespace OfficeIMO.Rtf;

/// <content>
/// Provides asynchronous lossless save APIs for read results.
/// </content>
public sealed partial class RtfReadResult {
    /// <summary>
    /// Saves the original RTF stream to a file without semantic normalization.
    /// </summary>
    public Task SaveLosslessAsync(string path, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytesAsync(
            path, ToBytesLossless(), cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Saves the original RTF stream to a stream without semantic normalization.
    /// </summary>
    public Task SaveLosslessAsync(Stream stream, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return OfficeIMO.Core.Internal.OfficeStreamWriter.WriteAllBytesAsync(
            stream, ToBytesLossless(), cancellationToken);
    }
}
