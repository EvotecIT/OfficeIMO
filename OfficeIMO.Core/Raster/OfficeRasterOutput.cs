using System;
using System.IO;

namespace OfficeIMO.Drawing;

internal static class OfficeRasterOutput {
    internal static void EnsureWritable(Stream destination) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
    }

    internal static bool TryGetMemoryStream(Stream destination, out MemoryStream? memoryStream) {
        EnsureWritable(destination);
        while (destination is OfficeImageExportEncodingStream guarded) destination = guarded.WrappedDestination;
        memoryStream = destination as MemoryStream;
        return memoryStream != null;
    }

    /// <summary>Returns the caller's retained backing array, including an exposed larger segment owner.</summary>
    internal static long GetMemoryStreamBackingBytes(MemoryStream stream) {
        long capacity = stream.Capacity;
        return stream.TryGetBuffer(out ArraySegment<byte> segment) && segment.Array != null
            ? Math.Max(capacity, segment.Array.LongLength) : capacity;
    }

    /// <summary>Checks source, conversion scratch, header, and retained caller output before writing.</summary>
    internal static void EnsureImageWriteWorkingSet(OfficeRasterImage image, Stream destination,
        int encodedBytes, int scratchBytes, int headerBytes, string message) {
        long outputPeak = TryGetMemoryStream(destination, out MemoryStream? memory)
            ? GetMemoryStreamBlockWritePeakBytes(memory!, encodedBytes, false, scratchBytes, headerBytes, scratchBytes)
            : 0L;
        long peak = image.PixelBuffer.LongLength + 24L + outputPeak + scratchBytes + 24L + headerBytes + 24L;
        if (peak > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException(message, nameof(image));
    }

    /// <summary>Bounds a complete append whose individual write lengths are not known.</summary>
    /// <remarks>
    /// A large write can seed any capacity, after which another write doubles it. Immediately before
    /// a growth the old capacity is below the final required length; use that envelope rather than
    /// assuming capacities are powers of two. No-growth caller buffers retain their exact size.
    /// </remarks>
    internal static long GetMemoryStreamWritePeakBytes(MemoryStream stream, long encodedBytes, bool materializeOutput) {
        long required = RequiredLength(stream, encodedBytes);
        long backing = GetMemoryStreamBackingBytes(stream);
        if (required <= stream.Capacity) return WithMaterializedCopy(backing + 24L, backing, required, materializeOutput);
        long capacity = Math.Max(256L, Math.Min(int.MaxValue, 2L * (required - 1L)));
        long peak = Math.Max(backing, required - 1L) + capacity + 48L;
        return WithMaterializedCopy(peak, capacity, required, materializeOutput);
    }

    /// <summary>Bounds one actual append, including the old and new arrays during a resize.</summary>
    internal static long GetMemoryStreamSingleWritePeakBytes(MemoryStream stream, long encodedBytes, bool materializeOutput) {
        long required = RequiredLength(stream, encodedBytes);
        long backing = GetMemoryStreamBackingBytes(stream);
        long capacity = stream.Capacity;
        long peak = backing + 24L;
        if (required > capacity) {
            capacity = GrowthCapacity(capacity, required);
            peak = backing + capacity + 48L;
            backing = capacity;
        }
        return WithMaterializedCopy(peak, backing, required, materializeOutput);
    }

    /// <summary>Bounds known prefix writes followed by writes no larger than one conversion block.</summary>
    /// <remarks>The prefix seeds an actual capacity; subsequent growth doubles once capacity covers a block.</remarks>
    internal static long GetMemoryStreamBlockWritePeakBytes(
        MemoryStream stream, long encodedBytes, bool materializeOutput, int maximumBlockBytes,
        int firstWriteBytes, int secondWriteBytes, int thirdWriteBytes = 0) {
        long required = RequiredLength(stream, encodedBytes);
        if (maximumBlockBytes < 1 || firstWriteBytes < 0 || secondWriteBytes < 0 || thirdWriteBytes < 0 ||
            firstWriteBytes + (long)secondWriteBytes + thirdWriteBytes > encodedBytes) {
            throw new ArgumentOutOfRangeException(nameof(maximumBlockBytes));
        }
        long backing = GetMemoryStreamBackingBytes(stream);
        long capacity = stream.Capacity;
        long position = stream.Position;
        long peak = backing + 24L;
        Append(firstWriteBytes); Append(secondWriteBytes); Append(thirdWriteBytes);
        if (capacity < maximumBlockBytes && capacity < required) {
            // This unusual prefix does not establish a doubling-only remainder. Keep the safe envelope.
            return Math.Max(peak, GetMemoryStreamWritePeakBytes(stream, encodedBytes, materializeOutput));
        }
        while (capacity < required) {
            long next = GrowthCapacity(capacity, Math.Min(required, capacity + 1L));
            peak = Math.Max(peak, backing + next + 48L);
            capacity = backing = next;
        }
        return WithMaterializedCopy(peak, backing, required, materializeOutput);

        void Append(int count) {
            position += count;
            if (position <= capacity) return;
            long next = GrowthCapacity(capacity, position);
            peak = Math.Max(peak, backing + next + 48L);
            capacity = backing = next;
        }
    }

    private static long RequiredLength(MemoryStream stream, long encodedBytes) {
        if (encodedBytes < 0L) throw new ArgumentOutOfRangeException(nameof(encodedBytes));
        if (encodedBytes > int.MaxValue - stream.Position) {
            throw new ArgumentException("Image encoding exceeds MemoryStream capacity limits.", nameof(stream));
        }
        return Math.Max(stream.Position + encodedBytes, stream.Length);
    }

    private static long GrowthCapacity(long capacity, long required) =>
        Math.Max(required, Math.Max(256L, Math.Min(int.MaxValue, capacity * 2L)));

    private static long WithMaterializedCopy(long peak, long backing, long length, bool materialize) =>
        materialize ? Math.Max(peak, backing + length + 48L) : peak;
}
