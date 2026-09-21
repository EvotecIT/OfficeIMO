using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents one serialized indirect object without requiring a retained stream payload to be
/// copied into another object-sized buffer before final file assembly.
/// </summary>
internal sealed class PdfSerializedObject {
    private readonly byte[]? _bytes;
    private readonly byte[]? _prefix;
    private readonly PdfStream? _stream;
    private readonly byte[]? _suffix;

    private PdfSerializedObject(byte[] bytes) {
        _bytes = bytes;
        Length = bytes.LongLength;
    }

    private PdfSerializedObject(byte[] prefix, PdfStream stream, byte[] suffix) {
        _prefix = prefix;
        _stream = stream;
        _suffix = suffix;
        Length = checked(prefix.LongLength + stream.DataLongLength + suffix.LongLength);
    }

    internal long Length { get; }

    internal static PdfSerializedObject FromBytes(byte[] bytes) {
        Guard.NotNull(bytes, nameof(bytes));
        return new PdfSerializedObject(bytes);
    }

    internal static PdfSerializedObject FromStream(byte[] prefix, PdfStream stream, byte[] suffix) {
        Guard.NotNull(prefix, nameof(prefix));
        Guard.NotNull(stream, nameof(stream));
        Guard.NotNull(suffix, nameof(suffix));
        return new PdfSerializedObject(prefix, stream, suffix);
    }

    internal void CopyTo(
        Stream destination,
        HashAlgorithm? hash,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(destination, nameof(destination));
        if (_bytes is not null) {
            CopySegment(_bytes, 0, _bytes.Length, destination, hash, cancellationToken);
            return;
        }

        CopySegment(_prefix!, 0, _prefix!.Length, destination, hash, cancellationToken);
        _stream!.GetDataSegment(out byte[] buffer, out int offset, out int length);
        CopySegment(buffer, offset, length, destination, hash, cancellationToken);
        CopySegment(_suffix!, 0, _suffix!.Length, destination, hash, cancellationToken);
    }

    private static void CopySegment(
        byte[] buffer,
        int offset,
        int length,
        Stream destination,
        HashAlgorithm? hash,
        CancellationToken cancellationToken) {
        if (!cancellationToken.CanBeCanceled) {
            hash?.TransformBlock(buffer, offset, length, buffer, offset);
            destination.Write(buffer, offset, length);
            return;
        }

        const int ChunkSize = 81920;
        int position = offset;
        int remaining = length;
        while (remaining > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(ChunkSize, remaining);
            hash?.TransformBlock(buffer, position, count, buffer, position);
            destination.Write(buffer, position, count);
            position += count;
            remaining -= count;
        }
    }
}
