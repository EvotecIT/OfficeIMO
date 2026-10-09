#nullable enable

#if NET8_0_OR_GREATER
using System;
using System.Buffers;
using System.Data;
using System.Text;

namespace OfficeIMO.Data;

public static partial class DataReaderUtf8TextExtensions {
    private const int Utf8EncodingBufferSize = 4096;

    /// <summary>Copies a byte segment of the current field's UTF-8 text into a caller-owned buffer.</summary>
    /// <param name="record">The record containing the current row.</param>
    /// <param name="ordinal">The zero-based field ordinal.</param>
    /// <param name="dataOffset">The zero-based byte offset in the encoded text.</param>
    /// <param name="buffer">The destination, or null to query the total encoded byte length.</param>
    /// <param name="bufferOffset">The zero-based destination byte offset.</param>
    /// <param name="length">The maximum number of bytes to copy.</param>
    /// <returns>The number of copied bytes, or the total UTF-8 byte length when buffer is null.</returns>
    /// <remarks>
    /// This method first uses <see cref="IDataReaderUtf8TextSource"/> when available. Otherwise it
    /// encodes <see cref="IDataRecord.GetString(int)"/> with the standard UTF-8 replacement fallback
    /// and bounded pooled storage. The provider's normal string conversion and cursor rules apply.
    /// Offsets and lengths are bytes, so a segment may split a multi-byte character. An offset at or
    /// beyond the end returns zero. All offsets and lengths must be non-negative, including length
    /// queries. A null buffer ignores dataOffset when returning the total length.
    /// </remarks>
    /// <exception cref="ArgumentNullException">The record is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">An offset or length is negative, or bufferOffset exceeds the buffer length.</exception>
    /// <exception cref="ArgumentException">The requested destination segment extends beyond the buffer.</exception>
    /// <exception cref="IndexOutOfRangeException">The ordinal is outside the record's fields.</exception>
    public static long GetUtf8Bytes(
        this IDataRecord record,
        int ordinal,
        long dataOffset,
        byte[]? buffer,
        int bufferOffset,
        int length) {
        if (record == null) throw new ArgumentNullException(nameof(record));
        if (dataOffset < 0) throw new ArgumentOutOfRangeException(nameof(dataOffset));
        if (bufferOffset < 0) throw new ArgumentOutOfRangeException(nameof(bufferOffset));
        if (length < 0) throw new ArgumentOutOfRangeException(nameof(length));
        if (buffer != null) {
            if (bufferOffset > buffer.Length) throw new ArgumentOutOfRangeException(nameof(bufferOffset));
            if (length > buffer.Length - bufferOffset) {
                throw new ArgumentException("The requested destination segment extends beyond the buffer.", nameof(length));
            }
        }

        if (record.TryGetUtf8Text(ordinal, out ReadOnlySpan<byte> borrowed)) {
            if (buffer == null) return borrowed.Length;
            if (dataOffset >= borrowed.Length) return 0;
            int offset = (int)dataOffset;
            int count = Math.Min(length, borrowed.Length - offset);
            borrowed.Slice(offset, count).CopyTo(buffer.AsSpan(bufferOffset, count));
            return count;
        }

        string text = record.GetString(ordinal);
        int byteCount = Encoding.UTF8.GetByteCount(text);
        if (buffer == null) return byteCount;
        if (dataOffset >= byteCount || length == 0) return 0;
        int copyLength = Math.Min(length, byteCount - (int)dataOffset);
        return CopyEncodedText(text, dataOffset, buffer, bufferOffset, copyLength);
    }

    private static int CopyEncodedText(string text, long dataOffset, byte[] buffer, int bufferOffset, int copyLength) {
        byte[] temporary = ArrayPool<byte>.Shared.Rent(Utf8EncodingBufferSize);
        try {
            Encoder encoder = Encoding.UTF8.GetEncoder();
            int charOffset = 0;
            long encodedOffset = 0;
            int copied = 0;
            while (copied < copyLength) {
                encoder.Convert(
                    text.AsSpan(charOffset), temporary.AsSpan(0, Utf8EncodingBufferSize), flush: true,
                    out int charsUsed, out int bytesUsed, out _);
                charOffset += charsUsed;
                int skip = (int)Math.Min(bytesUsed, Math.Max(0, dataOffset - encodedOffset));
                int count = Math.Min(bytesUsed - skip, copyLength - copied);
                temporary.AsSpan(skip, count).CopyTo(buffer.AsSpan(bufferOffset + copied, count));
                copied += count;
                encodedOffset += bytesUsed;
            }
            return copied;
        } finally {
            ArrayPool<byte>.Shared.Return(temporary, clearArray: true);
        }
    }
}
#endif
