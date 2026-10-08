using System;
using System.IO;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterImageDecoder {
    /// <summary>Reads owned encoded bytes within the shared raster size and working-memory limits.</summary>
    /// <remarks>
    /// Reads from the current position, leaves the stream open, and restores a seekable stream's
    /// original position on success or failure. A nonseekable stream is consumed and may advance
    /// one byte beyond the size limit to detect excess input. Returned bytes never borrow a
    /// caller-owned memory buffer. This operation bounds input; it does not validate an image format.
    /// </remarks>
    /// <exception cref="InvalidDataException">The input is empty, incomplete, or exceeds the configured limits.</exception>
    public static byte[] ReadEncodedBytes(Stream stream, OfficeRasterDecodeOptions? options = null) {
        return ReadEncodedBytes(stream, out _, options);
    }

    // Carries known caller-owned backing storage into parsers using the owned result.
    internal static byte[] ReadEncodedBytes(Stream stream, out long additionallyRetainedBytes, OfficeRasterDecodeOptions? options = null) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }
        additionallyRetainedBytes = 0L;
        var effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        long originalPosition = stream.CanSeek ? stream.Position : 0L;
        try {
            if (!OfficeBoundedStreamReader.TryRead(stream, effective.MaximumEncodedBytes,
                    effective.CancellationToken, out byte[] bytes, out additionallyRetainedBytes)) {
                throw new InvalidDataException("Encoded image input is empty, incomplete, or exceeds the configured limits.");
            }
            if (stream is MemoryStream memory && memory.TryGetBuffer(out ArraySegment<byte> buffer) &&
                ReferenceEquals(buffer.Array, bytes)) {
                if (!OfficeBoundedStreamReader.IsFinalCopyWithinLimit(bytes.LongLength, bytes.LongLength)) {
                    throw new InvalidDataException("Encoded image copy exceeds the managed working-memory limit.");
                }
                additionallyRetainedBytes = bytes.LongLength;
                var owned = new byte[bytes.Length];
                const int chunkSize = 64 * 1024;
                for (int offset = 0; offset < bytes.Length; offset += chunkSize) {
                    effective.CancellationToken.ThrowIfCancellationRequested();
                    Buffer.BlockCopy(bytes, offset, owned, offset, Math.Min(chunkSize, bytes.Length - offset));
                }
                return owned;
            }
            effective.CancellationToken.ThrowIfCancellationRequested();
            return bytes;
        } finally {
            if (stream.CanSeek) {
                stream.Position = originalPosition;
            }
        }
    }
}
