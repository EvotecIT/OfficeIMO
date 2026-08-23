using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static class OfficeBoundedStreamReader {
    internal static bool TryRead(
        Stream stream,
        int maximumBytes,
        CancellationToken cancellationToken,
        out byte[] bytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The source stream must be readable.", nameof(stream));
        if (maximumBytes < 1 || maximumBytes > OfficeRasterGuards.MaximumEncodedBytes) {
            throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        }

        bytes = Array.Empty<byte>();
        cancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek) {
            long remaining = stream.Length - stream.Position;
            if (remaining <= 0L || remaining > maximumBytes || remaining > int.MaxValue) return false;
            bytes = new byte[(int)remaining];
            return TryReadExact(stream, bytes, cancellationToken);
        }

        using var buffer = new MemoryStream(Math.Min(maximumBytes, 64 * 1024));
        var chunk = new byte[16 * 1024];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int remaining = maximumBytes - checked((int)buffer.Length);
            int read = stream.Read(chunk, 0, Math.Min(chunk.Length, remaining + 1));
            if (read <= 0) break;
            if (read > remaining) return false;
            buffer.Write(chunk, 0, read);
        }
        if (buffer.Length == 0L) return false;
        bytes = buffer.ToArray();
        return true;
    }

    private static bool TryReadExact(Stream stream, byte[] bytes, CancellationToken cancellationToken) {
        int offset = 0;
        while (offset < bytes.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(bytes, offset, bytes.Length - offset);
            if (read <= 0) return false;
            offset += read;
        }
        return true;
    }
}

