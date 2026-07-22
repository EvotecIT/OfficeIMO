using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Snapshots complete Office artifacts from caller-owned streams without changing seekable stream state.
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficeStreamReader {
        private const int BufferSize = 81920;

        /// <summary>
        /// Reads a complete artifact. Seekable sources are read from the beginning and restored to their original
        /// position; non-seekable sources are read forward from their current position.
        /// </summary>
        public static byte[] ReadAllBytes(Stream source, long? maxBytes = null) =>
            ReadAllBytes(source, CancellationToken.None, maxBytes);

        /// <summary>
        /// Reads from the caller's current position through the end of the stream without rewinding.
        /// </summary>
        public static byte[] ReadRemainingBytes(Stream source, long? maxBytes = null) {
            return ReadRemainingBytes(source, CancellationToken.None, maxBytes);
        }

        /// <summary>
        /// Reads from the caller's current position through the end of the stream without rewinding,
        /// with cooperative cancellation.
        /// </summary>
        public static byte[] ReadRemainingBytes(
            Stream source,
            CancellationToken cancellationToken,
            long? maxBytes = null) {
            ValidateRemainingSource(source, maxBytes);
            return ReadToEnd(source, cancellationToken, maxBytes);
        }

        /// <summary>Reads a complete artifact with cooperative cancellation.</summary>
        public static byte[] ReadAllBytes(
            Stream source,
            CancellationToken cancellationToken,
            long? maxBytes = null) {
            ValidateSource(source, maxBytes);
            long originalPosition = source.CanSeek ? source.Position : 0;
            try {
                if (source.CanSeek) source.Seek(0, SeekOrigin.Begin);
                return ReadToEnd(source, cancellationToken, maxBytes);
            } finally {
                if (source.CanSeek) source.Seek(originalPosition, SeekOrigin.Begin);
            }
        }

        /// <summary>
        /// Asynchronously reads a complete artifact while preserving the position of seekable caller-owned streams.
        /// </summary>
        public static async Task<byte[]> ReadAllBytesAsync(
            Stream source,
            CancellationToken cancellationToken = default,
            long? maxBytes = null) {
            ValidateSource(source, maxBytes);
            long originalPosition = source.CanSeek ? source.Position : 0;
            try {
                if (source.CanSeek) source.Seek(0, SeekOrigin.Begin);
                using var output = new MemoryStream();
                var buffer = new byte[BufferSize];
                long total = 0;
                int read;
                while ((read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0) {
                    total = checked(total + read);
                    EnsureWithinLimit(total, maxBytes);
                    await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                }
                return output.ToArray();
            } finally {
                if (source.CanSeek) source.Seek(originalPosition, SeekOrigin.Begin);
            }
        }

        private static byte[] ReadToEnd(Stream source, CancellationToken cancellationToken, long? maxBytes) {
            using var output = new MemoryStream();
            var buffer = new byte[BufferSize];
            long total = 0;
            int read;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total = checked(total + read);
                EnsureWithinLimit(total, maxBytes);
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }

        private static void ValidateSource(Stream source, long? maxBytes) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!source.CanRead) throw new ArgumentException("Stream must be readable.", nameof(source));
            if (maxBytes.HasValue && maxBytes.Value < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            if (source.CanSeek && maxBytes.HasValue && source.Length > maxBytes.Value) {
                throw new InvalidDataException($"Stream exceeds the configured maximum size ({maxBytes.Value} bytes).");
            }
        }

        private static void ValidateRemainingSource(Stream source, long? maxBytes) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!source.CanRead) throw new ArgumentException("Stream must be readable.", nameof(source));
            if (maxBytes.HasValue && maxBytes.Value < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            if (source.CanSeek && maxBytes.HasValue && source.Length - source.Position > maxBytes.Value) {
                throw new InvalidDataException($"Remaining stream content exceeds the configured maximum size ({maxBytes.Value} bytes).");
            }
        }

        private static void EnsureWithinLimit(long total, long? maxBytes) {
            if (maxBytes.HasValue && total > maxBytes.Value) {
                throw new InvalidDataException($"Stream exceeds the configured maximum size ({maxBytes.Value} bytes).");
            }
        }
    }
}
