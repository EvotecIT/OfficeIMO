using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Writes complete Office artifacts to caller-owned streams using one consistent overwrite contract.
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficeStreamWriter {
        /// <summary>
        /// Returns whether a stream can safely act as a document's associated destination. Associated destinations
        /// must support replacing the complete artifact on every parameterless save, not merely appending bytes at
        /// the stream's current position.
        /// </summary>
        public static bool CanReplaceContents(Stream destination) =>
            destination != null && destination.CanWrite && destination.CanSeek;

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
        /// Writes a complete artifact directly without closing the destination. Seekable streams are positioned
        /// at the beginning before production, then truncated to the completed artifact and rewound only after the
        /// producer succeeds. This preserves existing contents when validation fails before output begins.
        /// </summary>
        public static void Write(Stream destination, Action<Stream> writer) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(destination);
            ArgumentNullException.ThrowIfNull(writer);
#else
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (writer == null) throw new ArgumentNullException(nameof(writer));
#endif
            EnsureWritable(destination);
            long originalPosition = PositionForDirectWrite(destination);
            try {
                writer(destination);
                CompleteDirectWrite(destination);
                destination.Flush();
                RewindDestination(destination);
            } catch {
                RestorePositionAfterFailedDirectWrite(destination, originalPosition);
                throw;
            }
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

        /// <summary>
        /// Writes a complete artifact directly and asynchronously flushes it without closing the destination.
        /// The producer runs synchronously so format writers can preserve ordered offset accounting. Seekable
        /// streams are truncated only after the producer succeeds.
        /// </summary>
        public static async Task WriteAsync(
            Stream destination,
            Action<Stream> writer,
            CancellationToken cancellationToken = default) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(destination);
            ArgumentNullException.ThrowIfNull(writer);
#else
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (writer == null) throw new ArgumentNullException(nameof(writer));
#endif
            EnsureWritable(destination);
            cancellationToken.ThrowIfCancellationRequested();
            long originalPosition = PositionForDirectWrite(destination);
            try {
                writer(destination);
                cancellationToken.ThrowIfCancellationRequested();
                CompleteDirectWrite(destination);
                await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
                RewindDestination(destination);
            } catch {
                RestorePositionAfterFailedDirectWrite(destination, originalPosition);
                throw;
            }
        }

        /// <summary>
        /// Produces a complete artifact with asynchronous I/O without closing the destination. Seekable streams are
        /// truncated to the completed artifact and rewound after the producer succeeds.
        /// </summary>
        public static async Task WriteAsync(
            Stream destination,
            Func<Stream, CancellationToken, Task> writer,
            CancellationToken cancellationToken = default) {
#if NET6_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(destination);
            ArgumentNullException.ThrowIfNull(writer);
#else
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (writer == null) throw new ArgumentNullException(nameof(writer));
#endif
            EnsureWritable(destination);
            cancellationToken.ThrowIfCancellationRequested();
            long originalPosition = PositionForDirectWrite(destination);
            try {
                await writer(destination, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                CompleteDirectWrite(destination);
                await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
                RewindDestination(destination);
            } catch {
                RestorePositionAfterFailedDirectWrite(destination, originalPosition);
                throw;
            }
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

        private static long PositionForDirectWrite(Stream destination) {
            if (!destination.CanSeek) return 0L;
            long originalPosition = destination.Position;
            destination.Position = 0;
            return originalPosition;
        }

        private static void CompleteDirectWrite(Stream destination) {
            if (!destination.CanSeek) return;
            destination.SetLength(destination.Position);
        }

        private static void RestorePositionAfterFailedDirectWrite(Stream destination, long originalPosition) {
            if (!destination.CanSeek) return;
            destination.Position = Math.Min(originalPosition, destination.Length);
        }

        private static void RewindDestination(Stream destination) {
            if (destination.CanSeek) destination.Position = 0;
        }
    }
}
