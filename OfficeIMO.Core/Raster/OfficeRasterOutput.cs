using System;
using System.IO;

namespace OfficeIMO.Drawing;

internal static class OfficeRasterOutput {
    internal static void EnsureWritable(Stream destination) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) {
            throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        }
    }

    internal static bool TryGetMemoryStream(Stream destination, out MemoryStream? memoryStream) {
        EnsureWritable(destination);
        while (destination is OfficeImageExportEncodingStream guarded) {
            destination = guarded.WrappedDestination;
        }
        memoryStream = destination as MemoryStream;
        return memoryStream != null;
    }

    /// <summary>Returns the caller's retained backing array, including an exposed larger segment owner.</summary>
    internal static long GetMemoryStreamBackingBytes(MemoryStream stream) {
        long capacity = stream.Capacity;
        return stream.TryGetBuffer(out ArraySegment<byte> segment) && segment.Array != null
            ? Math.Max(capacity, segment.Array.LongLength)
            : capacity;
    }

    /// <summary>Bounds retained output and transient backing-array growth for a complete append.</summary>
    /// <remarks>
    /// Incremental writes can double capacity repeatedly. Project each doubling through the final
    /// position rather than assuming that one allocation has the exact encoded-image size.
    /// Caller streams do not require a final array copy; byte-array materialization does.
    /// </remarks>
    internal static long GetMemoryStreamWritePeakBytes(
        MemoryStream stream, long encodedBytes, bool materializeOutput) {
        if (encodedBytes < 0L) throw new ArgumentOutOfRangeException(nameof(encodedBytes));
        try {
            long requiredLength = Math.Max(checked(stream.Position + encodedBytes), stream.Length);
            long backingBytes = GetMemoryStreamBackingBytes(stream);
            long capacity = stream.Capacity;
            long peakBytes = checked(backingBytes + 24L);
            while (capacity < requiredLength) {
                long projectedCapacity = Math.Max(256L, checked(capacity * 2L));
                peakBytes = Math.Max(peakBytes, checked(backingBytes + 24L + projectedCapacity + 24L));
                backingBytes = projectedCapacity;
                capacity = projectedCapacity;
            }
            if (materializeOutput) {
                peakBytes = Math.Max(peakBytes, checked(backingBytes + 24L + requiredLength + 24L));
            }
            return peakBytes;
        } catch (OverflowException) {
            throw new ArgumentException("Image encoding exceeds the managed working-set limit.", nameof(stream));
        }
    }
}
