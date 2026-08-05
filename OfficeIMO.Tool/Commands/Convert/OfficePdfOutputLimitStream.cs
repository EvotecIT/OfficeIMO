namespace OfficeIMO.Tool.Commands.Convert;

/// <summary>Rejects PDF output before it exceeds the configured byte limit.</summary>
internal sealed class OfficePdfOutputLimitStream : Stream {
    private readonly Stream _destination;
    private readonly long _maximumLength;

    internal OfficePdfOutputLimitStream(Stream destination, long maximumLength) {
        _destination = destination ?? throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        if (maximumLength <= 0) throw new ArgumentOutOfRangeException(nameof(maximumLength));
        _maximumLength = maximumLength;
    }

    public override bool CanRead => false;
    public override bool CanSeek => _destination.CanSeek;
    public override bool CanWrite => true;
    public override long Length => _destination.Length;

    public override long Position {
        get => _destination.Position;
        set {
            EnsureWithinLimit(value);
            _destination.Position = value;
        }
    }

    public override void Flush() => _destination.Flush();

    public override Task FlushAsync(CancellationToken cancellationToken) =>
        _destination.FlushAsync(cancellationToken);

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) {
        if (!_destination.CanSeek) throw new NotSupportedException();
        long originPosition = origin switch {
            SeekOrigin.Begin => 0L,
            SeekOrigin.Current => _destination.Position,
            SeekOrigin.End => _destination.Length,
            _ => throw new ArgumentOutOfRangeException(nameof(origin))
        };
        long target = checked(originPosition + offset);
        EnsureWithinLimit(target);
        return _destination.Seek(offset, origin);
    }

    public override void SetLength(long value) {
        EnsureWithinLimit(value);
        _destination.SetLength(value);
    }

    public override void Write(byte[] buffer, int offset, int count) {
        ArgumentNullException.ThrowIfNull(buffer);
        ValidateBufferRange(buffer, offset, count);
        long end = checked(Position + count);
        EnsureWithinLimit(end);
        _destination.Write(buffer, offset, count);
    }

    public override void Write(ReadOnlySpan<byte> buffer) {
        long end = checked(Position + buffer.Length);
        EnsureWithinLimit(end);
        _destination.Write(buffer);
    }

    public override void WriteByte(byte value) {
        long end = checked(Position + 1L);
        EnsureWithinLimit(end);
        _destination.WriteByte(value);
    }

    public override Task WriteAsync(
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(buffer);
        ValidateBufferRange(buffer, offset, count);
        long end = checked(Position + count);
        EnsureWithinLimit(end);
        return _destination.WriteAsync(buffer, offset, count, cancellationToken);
    }

    public override ValueTask WriteAsync(
        ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken = default) {
        long end = checked(Position + buffer.Length);
        EnsureWithinLimit(end);
        return _destination.WriteAsync(buffer, cancellationToken);
    }

    protected override void Dispose(bool disposing) {
        // The command owns the destination stream.
        base.Dispose(disposing);
    }

    private void EnsureWithinLimit(long length) {
        if (length < 0 || length > _maximumLength) {
            throw new OfficePdfOutputLimitException(length, _maximumLength);
        }
    }

    private static void ValidateBufferRange(byte[] buffer, int offset, int count) {
        if ((uint)offset > (uint)buffer.Length) throw new ArgumentOutOfRangeException(nameof(offset));
        if ((uint)count > (uint)(buffer.Length - offset)) throw new ArgumentOutOfRangeException(nameof(count));
    }
}

internal sealed class OfficePdfOutputLimitException : IOException {
    internal OfficePdfOutputLimitException(long observedBytes, long maximumBytes)
        : base("PDF output size " + observedBytes + " bytes exceeds the configured maximum of " + maximumBytes + " bytes.") {
        ObservedBytes = observedBytes;
        MaximumBytes = maximumBytes;
    }

    internal long ObservedBytes { get; }
    internal long MaximumBytes { get; }
}
