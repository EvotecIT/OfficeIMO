namespace OfficeIMO.Workflows;

/// <summary>
/// Seekable in-memory destination that fails before a workflow artifact can grow beyond its configured output budget.
/// </summary>
internal sealed class OfficeWorkflowBoundedMemoryStream : MemoryStream {
    private readonly long _maximumBytes;

    internal OfficeWorkflowBoundedMemoryStream(long maximumBytes) {
        if (maximumBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
    }

    public override void SetLength(long value) {
        EnsureWithinLimit(value);
        base.SetLength(value);
    }

    public override void Write(byte[] buffer, int offset, int count) {
        EnsureWriteWithinLimit(count);
        base.Write(buffer, offset, count);
    }

    public override void Write(ReadOnlySpan<byte> buffer) {
        EnsureWriteWithinLimit(buffer.Length);
        base.Write(buffer);
    }

    public override void WriteByte(byte value) {
        EnsureWriteWithinLimit(1);
        base.WriteByte(value);
    }

    public override Task WriteAsync(
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        EnsureWriteWithinLimit(count);
        return base.WriteAsync(buffer, offset, count, cancellationToken);
    }

    public override ValueTask WriteAsync(
        ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken = default) {
        EnsureWriteWithinLimit(buffer.Length);
        return base.WriteAsync(buffer, cancellationToken);
    }

    private void EnsureWriteWithinLimit(int count) {
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        long end;
        try {
            end = checked(Position + count);
        } catch (OverflowException) {
            throw CreateLimitException();
        }
        EnsureWithinLimit(Math.Max(Length, end));
    }

    private void EnsureWithinLimit(long length) {
        if (length > _maximumBytes) throw CreateLimitException();
    }

    private InvalidOperationException CreateLimitException() => new(
        $"Generated artifact exceeded the configured {_maximumBytes:N0}-byte output limit while it was being serialized.");
}
