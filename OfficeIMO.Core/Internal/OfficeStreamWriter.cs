using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
    /// <summary>
    /// Writes complete Office artifacts to caller-owned streams using one consistent overwrite contract.
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficeStreamWriter {
        /// <summary>
        /// Writes a complete artifact without closing the destination. Seekable streams are truncated before
        /// writing and rewound after a successful write; non-seekable streams are written at their current position.
        /// </summary>
        public static void WriteAllBytes(Stream destination, byte[] bytes) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(destination);
            ArgumentNullException.ThrowIfNull(bytes);
#else
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
#endif
            EnsureWritable(destination);
            PrepareDestination(destination);
            destination.Write(bytes, 0, bytes.Length);
            destination.Flush();
            RewindDestination(destination);
        }

        /// <summary>
        /// Asynchronously writes a complete artifact without closing the destination. Seekable streams are
        /// truncated before writing and rewound after a successful write.
        /// </summary>
        public static async Task WriteAllBytesAsync(
            Stream destination,
            byte[] bytes,
            CancellationToken cancellationToken = default) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(destination);
            ArgumentNullException.ThrowIfNull(bytes);
#else
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
#endif
            EnsureWritable(destination);
            cancellationToken.ThrowIfCancellationRequested();
            PrepareDestination(destination);
#if NET6_0_OR_GREATER
            await destination.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
            await destination.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
            await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
            RewindDestination(destination);
        }

        private static void EnsureWritable(Stream destination) {
            if (!destination.CanWrite) {
                throw new ArgumentException("Destination stream must be writable.", nameof(destination));
            }
        }

        private static void PrepareDestination(Stream destination) {
            if (!destination.CanSeek) return;
            destination.Position = 0;
            destination.SetLength(0);
        }

        private static void RewindDestination(Stream destination) {
            if (destination.CanSeek) destination.Position = 0;
        }
    }
}
