using Avalonia.Platform.Storage;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Consumes picker-owned input streams without treating provider URIs as local paths.</summary>
internal static class StudioStorageInput {
    internal const long MaximumImageBytes = 64L * 1024 * 1024;

    /// <summary>Reads the selected image and releases every returned picker item, including on cancellation.</summary>
    internal static async Task<byte[]?> ReadImageAsync(
        IReadOnlyList<IStorageFile> files,
        CancellationToken cancellationToken,
        long maximumBytes = MaximumImageBytes) {
        ArgumentNullException.ThrowIfNull(files);
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (files.Count == 0) return null;
            await using Stream input = await files[0].OpenReadAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(
                input, cancellationToken, maximumBytes).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return bytes;
        } finally {
            foreach (IStorageFile file in files.Distinct<IStorageFile>(ReferenceEqualityComparer.Instance)) {
                file.Dispose();
            }
        }
    }
}
