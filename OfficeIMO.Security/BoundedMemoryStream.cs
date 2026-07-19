using System.IO;

namespace OfficeIMO.Security;

/// <summary>Collects decoded security content without allowing the buffer to grow past a caller-owned limit.</summary>
internal sealed class BoundedMemoryStream : MemoryStream {
    private readonly long _maximumBytes;

    internal BoundedMemoryStream(long maximumBytes) {
        if (maximumBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maximumBytes), "The byte limit must be positive.");
        }
        _maximumBytes = maximumBytes;
    }

    public override void Write(byte[] buffer, int offset, int count) {
        EnsureWriteFits(count);
        base.Write(buffer, offset, count);
    }

    public override void WriteByte(byte value) {
        EnsureWriteFits(1);
        base.WriteByte(value);
    }

    public override void SetLength(long value) {
        if (value > _maximumBytes) throw CreateLimitException(value);
        base.SetLength(value);
    }

    private void EnsureWriteFits(int count) {
        long candidateLength = Position > long.MaxValue - count
            ? long.MaxValue
            : Math.Max(Length, Position + count);
        if (candidateLength > _maximumBytes) throw CreateLimitException(candidateLength);
    }

    private SecurityContentLimitExceededException CreateLimitException(long attemptedBytes) =>
        new(attemptedBytes, _maximumBytes);
}

internal sealed class SecurityContentLimitExceededException : IOException {
    internal SecurityContentLimitExceededException(long attemptedBytes, long maximumBytes)
        : base($"The decoded value would require {attemptedBytes} bytes and exceeds the configured limit of {maximumBytes} bytes.") { }
}
