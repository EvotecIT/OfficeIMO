namespace OfficeIMO.Core.Internal;

internal static partial class OfficeDeflateStreamValidator {
    /// <summary>Validates the remaining seekable stream with bounded buffering and restores its position.</summary>
    internal static bool TryValidateExact(Stream source, long maximumOutputBytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        if (!source.CanRead || !source.CanSeek || maximumOutputBytes < 0) return false;
        long position = source.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            long count = source.Length - position;
            if (count < 0) return false;
            return TryValidateExact(new DeflateBitReader(source, count), count, maximumOutputBytes, out _, cancellationToken);
        } finally {
            source.Position = position;
        }
    }

    /// <summary>Validates one complete Deflate stream that starts at the current position of a seekable stream and may be
    /// followed by other data, as in gzip members. Reports its compressed length and decoded size, then restores the position.</summary>
    internal static bool TryMeasure(Stream source, long count, long maximumOutputBytes, out long consumedBytes, out long outputBytes,
        out bool outputLimitExceeded, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        consumedBytes = 0; outputBytes = 0; outputLimitExceeded = false;
        if (!source.CanRead || !source.CanSeek || maximumOutputBytes < 0 || count < 0) return false;
        long position = source.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            count = Math.Min(count, source.Length - position);
            var reader = new DeflateBitReader(source, count);
            if (!TryValidateStream(reader, maximumOutputBytes, out outputLimitExceeded, out outputBytes, cancellationToken)) return false;
            consumedBytes = reader.ConsumedBytes;
            return true;
        } finally {
            source.Position = position;
        }
    }
}
