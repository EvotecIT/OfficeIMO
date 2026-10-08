using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

/// <summary>Saves already encoded image bytes through a same-directory atomic file commit.</summary>
/// <remarks>
/// The supplied bytes are borrowed and are not modified, cloned, parsed, or format-validated.
/// Keep them unchanged until the operation completes. An existing destination is overwritten
/// only after staging succeeds. Unsupported atomic replacement fails without a non-atomic fallback.
/// File-system failures propagate; failed staging is cleaned up when the file system permits it.
/// Cancellation is observed while staging and immediately before commit. Once replacement begins,
/// it can complete even if cancellation is requested. This API does not promise power-loss durability.
/// </remarks>
public static class OfficeImageFileWriter {
    /// <summary>Stages encoded bytes and atomically creates or replaces the destination file.</summary>
    /// <param name="path">The destination file path. Missing parent directories are created.</param>
    /// <param name="encodedImage">Already encoded bytes, borrowed until the method returns.</param>
    /// <param name="cancellationToken">Cancels staging and prevents publication when observed before commit.</param>
    /// <exception cref="ArgumentNullException">The supplied byte array is null.</exception>
    public static void WriteAllBytes(string path, byte[] encodedImage, CancellationToken cancellationToken = default) {
        if (encodedImage == null) {
            throw new ArgumentNullException(nameof(encodedImage));
        }
        OfficeFileCommit.WriteAtomically(path, stream => {
            const int chunkSize = 64 * 1024;
            for (int offset = 0; offset < encodedImage.Length;) {
                cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(chunkSize, encodedImage.Length - offset);
                stream.Write(encodedImage, offset, count);
                offset += count;
            }
        }, cancellationToken);
    }

    /// <summary>Asynchronously stages encoded bytes and atomically creates or replaces the destination file.</summary>
    /// <param name="path">The destination file path. Missing parent directories are created.</param>
    /// <param name="encodedImage">Already encoded bytes, borrowed until the returned task completes.</param>
    /// <param name="cancellationToken">Cancels staging and prevents publication when observed before commit.</param>
    /// <exception cref="ArgumentNullException">The supplied byte array is null.</exception>
    public static Task WriteAllBytesAsync(string path, byte[] encodedImage, CancellationToken cancellationToken = default) {
        if (encodedImage == null) {
            throw new ArgumentNullException(nameof(encodedImage));
        }
        return OfficeFileCommit.WriteAtomicallyAsync(path, async (stream, token) => {
            const int chunkSize = 64 * 1024;
            for (int offset = 0; offset < encodedImage.Length;) {
                token.ThrowIfCancellationRequested();
                int count = Math.Min(chunkSize, encodedImage.Length - offset);
                await stream.WriteAsync(encodedImage, offset, count, token).ConfigureAwait(false);
                offset += count;
            }
        }, cancellationToken);
    }
}
