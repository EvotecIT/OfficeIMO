using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal {
    /// <summary>
    /// Snapshots complete Office artifacts from caller-owned streams without changing seekable stream state.
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficeStreamReader {
        private const int BufferSize = 81920;
        private static readonly object SizeLimitMarker = new object();

        internal static bool IsSizeLimitException(InvalidDataException exception) =>
            exception.Data.Contains(SizeLimitMarker);

        internal static long? GetExceededSizeLimit(InvalidDataException exception) =>
            exception.Data[SizeLimitMarker] is long limit ? limit : (long?)null;

        private static InvalidDataException SizeLimitExceeded(string message, long limit) {
            var exception = new InvalidDataException(message);
            exception.Data[SizeLimitMarker] = limit;
            return exception;
        }

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
            long? maxBytes = null, int probeLength = 0, Func<byte[], int, long?>? probeLimit = null) {
            ValidateRemainingSource(source, maxBytes);
            return ReadToEnd(source, cancellationToken, maxBytes, probeLength, probeLimit);
        }

        /// <summary>Reads a complete artifact with cooperative cancellation.</summary>
        public static byte[] ReadAllBytes(
            Stream source,
            CancellationToken cancellationToken,
            long? maxBytes = null, int probeLength = 0, Func<byte[], int, long?>? probeLimit = null) {
            ValidateSource(source, maxBytes);
            long originalPosition = source.CanSeek ? source.Position : 0;
            try {
                if (source.CanSeek) source.Seek(0, SeekOrigin.Begin);
                return ReadToEnd(source, cancellationToken, maxBytes, probeLength, probeLimit);
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
            long? maxBytes = null, int probeLength = 0, Func<byte[], int, long?>? probeLimit = null) {
            ValidateSource(source, maxBytes);
            long originalPosition = source.CanSeek ? source.Position : 0;
            try {
                if (source.CanSeek) source.Seek(0, SeekOrigin.Begin);
                return await ReadToEndAsync(source, cancellationToken, maxBytes, probeLength, probeLimit).ConfigureAwait(false);
            } finally {
                if (source.CanSeek) source.Seek(originalPosition, SeekOrigin.Begin);
            }
        }

        /// <summary>
        /// Asynchronously reads from the caller's current position through the end of the stream without rewinding.
        /// </summary>
        public static Task<byte[]> ReadRemainingBytesAsync(
            Stream source,
            CancellationToken cancellationToken = default,
            long? maxBytes = null, int probeLength = 0, Func<byte[], int, long?>? probeLimit = null) {
            ValidateRemainingSource(source, maxBytes);
            return ReadToEndAsync(source, cancellationToken, maxBytes, probeLength, probeLimit);
        }

        private static async Task<byte[]> ReadToEndAsync(
            Stream source,
            CancellationToken cancellationToken,
            long? maxBytes, int probeLength, Func<byte[], int, long?>? probeLimit) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] prefix = Array.Empty<byte>();
            int prefixCount = 0;
            if (probeLimit != null) {
                prefix = new byte[GetProbeLength(probeLength, maxBytes)];
                long start = source.CanSeek ? source.Position : 0;
                while (prefixCount < prefix.Length) {
                    int count = await source.ReadAsync(prefix, prefixCount, prefix.Length - prefixCount, cancellationToken).ConfigureAwait(false);
                    if (count == 0) break;
                    prefixCount += count;
                }
                maxBytes = CombineLimits(maxBytes, probeLimit(prefix, prefixCount));
                if (source.CanSeek) source.Position = start;
                else EnsureWithinLimit(prefixCount, maxBytes);
            }
            if (source.CanSeek) {
                long remaining = Math.Max(0, source.Length - source.Position);
                EnsureWithinLimit(remaining, maxBytes);
                if (remaining > int.MaxValue) {
                    throw new InvalidDataException($"Stream length {remaining} exceeds the supported in-memory size.");
                }

                byte[] result = remaining == 0
                    ? Array.Empty<byte>()
                    : new byte[checked((int)remaining)];
                int offset = 0;
                while (offset < result.Length) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int exactRead = await source.ReadAsync(
                        result,
                        offset,
                        result.Length - offset,
                        cancellationToken).ConfigureAwait(false);
                    if (exactRead == 0) {
                        Array.Resize(ref result, offset);
                        return result;
                    }

                    offset += exactRead;
                }

                var trailingByte = new byte[1];
                int trailingRead = await source.ReadAsync(
                    trailingByte,
                    0,
                    1,
                    cancellationToken).ConfigureAwait(false);
                if (trailingRead == 0) return result;

                long expandedTotal = checked((long)result.Length + 1);
                EnsureWithinLimit(expandedTotal, maxBytes);
                using var expanded = new MemoryStream(checked(result.Length + 1));
                expanded.Write(result, 0, result.Length);
                expanded.WriteByte(trailingByte[0]);
                var expandedBuffer = new byte[BufferSize];
                int expandedRead;
                while ((expandedRead = await source.ReadAsync(
                    expandedBuffer,
                    0,
                    GetReadSize(expandedTotal, maxBytes),
                    cancellationToken).ConfigureAwait(false)) > 0) {
                    expandedTotal = checked(expandedTotal + expandedRead);
                    EnsureWithinLimit(expandedTotal, maxBytes);
                    expanded.Write(expandedBuffer, 0, expandedRead);
                }
                return expanded.ToArray();
            }

            using var output = new MemoryStream();
            var buffer = new byte[BufferSize];
            output.Write(prefix, 0, prefixCount);
            long total = prefixCount;
            int read;
            while ((read = await source.ReadAsync(buffer, 0, GetReadSize(total, maxBytes), cancellationToken).ConfigureAwait(false)) > 0) {
                total = checked(total + read);
                EnsureWithinLimit(total, maxBytes);
                await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
            }
            return output.ToArray();
        }

        private static byte[] ReadToEnd(Stream source, CancellationToken cancellationToken, long? maxBytes,
            int probeLength, Func<byte[], int, long?>? probeLimit) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[]? prefix = null;
            int prefixCount = 0;
            if (probeLimit != null) {
                prefix = new byte[GetProbeLength(probeLength, maxBytes)];
                long start = source.CanSeek ? source.Position : 0;
                while (prefixCount < prefix.Length) {
                    int count = source.Read(prefix, prefixCount, prefix.Length - prefixCount);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (count == 0) break;
                    prefixCount += count;
                }
                maxBytes = CombineLimits(maxBytes, probeLimit(prefix, prefixCount));
                if (source.CanSeek) source.Position = start;
                else EnsureWithinLimit(prefixCount, maxBytes);
            }
            if (source.CanSeek) {
                long remaining = Math.Max(0, source.Length - source.Position);
                EnsureWithinLimit(remaining, maxBytes);
                if (remaining > int.MaxValue) {
                    throw new InvalidDataException($"Stream length {remaining} exceeds the supported in-memory size.");
                }

                byte[] result = remaining == 0
                    ? Array.Empty<byte>()
                    : new byte[checked((int)remaining)];
                int offset = 0;
                while (offset < result.Length) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int exactRead = source.Read(result, offset, result.Length - offset);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (exactRead == 0) {
                        Array.Resize(ref result, offset);
                        return result;
                    }

                    offset += exactRead;
                }

                int trailingByte = source.ReadByte();
                cancellationToken.ThrowIfCancellationRequested();
                if (trailingByte < 0) return result;

                long expandedTotal = checked((long)result.Length + 1);
                EnsureWithinLimit(expandedTotal, maxBytes);
                using var expanded = new MemoryStream(checked(result.Length + 1));
                expanded.Write(result, 0, result.Length);
                expanded.WriteByte(checked((byte)trailingByte));
                var expandedBuffer = new byte[BufferSize];
                int expandedRead;
                while ((expandedRead = source.Read(expandedBuffer, 0, GetReadSize(expandedTotal, maxBytes))) > 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                    expandedTotal = checked(expandedTotal + expandedRead);
                    EnsureWithinLimit(expandedTotal, maxBytes);
                    expanded.Write(expandedBuffer, 0, expandedRead);
                }
                return expanded.ToArray();
            }

            return ReadForwardBytes(source, cancellationToken, maxBytes, prefix, prefixCount);
        }

        /// <summary>
        /// Reads forward without inspecting length or seeking, consuming at most one byte beyond a configured limit.
        /// Leaves caller-owned streams open, including when a limit or cancellation stops the read.
        /// </summary>
        internal static byte[] ReadForwardBytes(Stream source, CancellationToken cancellationToken, long? maxBytes = null, byte[]? prefix = null, int prefixCount = 0) {
            ArgumentNullException.ThrowIfNull(source);
            if (!source.CanRead) throw new ArgumentException("Stream must be readable.", nameof(source));
            if (maxBytes.HasValue && maxBytes.Value < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            using var output = new MemoryStream();
            var buffer = new byte[BufferSize];
            if (prefix != null) output.Write(prefix, 0, prefixCount);
            long total = prefixCount;
            int read;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                read = source.Read(buffer, 0, GetReadSize(total, maxBytes));
                cancellationToken.ThrowIfCancellationRequested();
                if (read == 0) break;
                total = checked(total + read);
                EnsureWithinLimit(total, maxBytes);
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }

        private static int GetProbeLength(int requested, long? maxBytes) =>
            maxBytes.HasValue && maxBytes.Value < requested ? (int)maxBytes.Value : requested;

        private static long? CombineLimits(long? original, long? detected) =>
            original.HasValue && detected.HasValue ? Math.Min(original.Value, detected.Value) : detected ?? original;

        private static int GetReadSize(long total, long? maxBytes) {
            if (!maxBytes.HasValue) return BufferSize;
            long remaining = maxBytes.Value - total;
            return remaining >= BufferSize ? BufferSize : checked((int)remaining + 1);
        }

        private static void ValidateSource(Stream source, long? maxBytes) {
            ArgumentNullException.ThrowIfNull(source);
            if (!source.CanRead) throw new ArgumentException("Stream must be readable.", nameof(source));
            if (maxBytes.HasValue && maxBytes.Value < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            if (source.CanSeek && maxBytes.HasValue && source.Length > maxBytes.Value) {
                throw SizeLimitExceeded($"Stream exceeds the configured maximum size ({maxBytes.Value} bytes).", maxBytes.Value);
            }
        }

        private static void ValidateRemainingSource(Stream source, long? maxBytes) {
            ArgumentNullException.ThrowIfNull(source);
            if (!source.CanRead) throw new ArgumentException("Stream must be readable.", nameof(source));
            if (maxBytes.HasValue && maxBytes.Value < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            if (source.CanSeek && maxBytes.HasValue && source.Length - source.Position > maxBytes.Value) {
                throw SizeLimitExceeded($"Remaining stream content exceeds the configured maximum size ({maxBytes.Value} bytes).", maxBytes.Value);
            }
        }

        private static void EnsureWithinLimit(long total, long? maxBytes) {
            if (maxBytes.HasValue && total > maxBytes.Value) {
                throw SizeLimitExceeded($"Stream exceeds the configured maximum size ({maxBytes.Value} bytes).", maxBytes.Value);
            }
        }
    }
}
