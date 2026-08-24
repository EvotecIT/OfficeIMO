using System;
using System.IO;
using System.Reflection;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static class OfficeBoundedStreamReader {
    private static readonly FieldInfo? MemoryStreamBufferField =
        typeof(MemoryStream).GetField("_buffer", BindingFlags.Instance | BindingFlags.NonPublic) ??
        typeof(MemoryStream).GetField("m_buffer", BindingFlags.Instance | BindingFlags.NonPublic);

    internal static bool TryRead(
        Stream stream,
        int maximumBytes,
        CancellationToken cancellationToken,
        out byte[] bytes) =>
        TryRead(stream, maximumBytes, cancellationToken, out bytes, out _);

    internal static bool TryRead(
        Stream stream,
        int maximumBytes,
        CancellationToken cancellationToken,
        out byte[] bytes,
        out long retainedManagedBytes) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The source stream must be readable.", nameof(stream));
        if (maximumBytes < 1 || maximumBytes > OfficeRasterGuards.MaximumEncodedBytes) {
            throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        }

        bytes = Array.Empty<byte>();
        retainedManagedBytes = 0L;
        cancellationToken.ThrowIfCancellationRequested();
        if (stream.CanSeek) {
            long remaining = stream.Length - stream.Position;
            if (remaining <= 0L || remaining > maximumBytes || remaining > int.MaxValue) return false;
            if (stream is MemoryStream memoryStream) {
                if (TryBorrowFullBuffer(memoryStream, remaining, out bytes)) return true;
                if (!TryGetRetainedMemoryStreamBytes(memoryStream, out long retainedStreamBytes) ||
                    !IsSeekableMemoryCopyWithinLimit(retainedStreamBytes, remaining)) return false;
                retainedManagedBytes = retainedStreamBytes;
            }
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
            int requiredCapacity = checked((int)buffer.Length + read);
            if (requiredCapacity > buffer.Capacity) {
                int doubledCapacity = checked(buffer.Capacity * 2);
                buffer.Capacity = Math.Min(maximumBytes, Math.Max(requiredCapacity, doubledCapacity));
            }
            buffer.Write(chunk, 0, read);
        }
        if (buffer.Length == 0L) return false;
        byte[] retainedBuffer = buffer.GetBuffer();
        int payloadLength = checked((int)buffer.Length);
        if (payloadLength == retainedBuffer.Length) {
            bytes = retainedBuffer;
            return true;
        }
        if (!IsFinalCopyWithinLimit(retainedBuffer.LongLength, payloadLength)) return false;
        bytes = new byte[payloadLength];
        CopyWithCancellation(retainedBuffer, bytes, cancellationToken);
        return true;
    }

    internal static bool IsFinalCopyWithinLimit(long retainedBufferBytes, long payloadBytes) {
        if (retainedBufferBytes < 0L || payloadBytes < 0L || payloadBytes > retainedBufferBytes) return false;
        try {
            return checked(retainedBufferBytes + 24L + payloadBytes + 24L) <=
                   OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    internal static bool IsSeekableMemoryCopyWithinLimit(long retainedStreamBytes, long payloadBytes) {
        if (retainedStreamBytes < 0L || payloadBytes < 0L || payloadBytes > retainedStreamBytes) return false;
        try {
            return checked(retainedStreamBytes + 24L + payloadBytes + 24L) <=
                   OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    internal static bool TryGetRetainedMemoryStreamBytes(
        MemoryStream stream,
        out long retainedStreamBytes) {
        if (stream.TryGetBuffer(out ArraySegment<byte> segment) && segment.Array != null) {
            retainedStreamBytes = segment.Array.LongLength;
            return true;
        }
        if (MemoryStreamBufferField?.GetValue(stream) is byte[] hiddenBuffer) {
            retainedStreamBytes = hiddenBuffer.LongLength;
            return true;
        }
        retainedStreamBytes = 0L;
        return stream.Capacity == 0;
    }

    private static bool TryBorrowFullBuffer(
        MemoryStream stream,
        long remaining,
        out byte[] bytes) {
        bytes = Array.Empty<byte>();
        if (stream.Position != 0L || remaining != stream.Length || remaining > int.MaxValue ||
            !stream.TryGetBuffer(out ArraySegment<byte> segment) || segment.Array == null ||
            segment.Offset != 0 || segment.Count != remaining || segment.Array.LongLength != remaining) {
            return false;
        }
        bytes = segment.Array;
        stream.Position = stream.Length;
        return true;
    }

    private static void CopyWithCancellation(
        byte[] source,
        byte[] destination,
        CancellationToken cancellationToken) {
        const int copyChunkBytes = 64 * 1024;
        int offset = 0;
        while (offset < destination.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(copyChunkBytes, destination.Length - offset);
            Buffer.BlockCopy(source, offset, destination, offset, count);
            offset += count;
        }
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

