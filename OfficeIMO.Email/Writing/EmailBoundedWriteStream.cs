namespace OfficeIMO.Email;

/// <summary>Counts produced bytes and rejects output before it exceeds the configured limit.</summary>
internal sealed class EmailBoundedWriteStream : Stream {
    private readonly Stream _destination;
    private readonly long _maximumLength;
    private long _maximumPosition;

    internal EmailBoundedWriteStream(Stream destination, long maximumLength) {
        _destination = destination ?? throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        if (maximumLength <= 0) throw new ArgumentOutOfRangeException(nameof(maximumLength));
        _maximumLength = maximumLength;
        _maximumPosition = destination.CanSeek ? destination.Position : 0L;
    }

    internal long BytesWritten => _maximumPosition;

    public override bool CanRead => false;
    public override bool CanSeek => _destination.CanSeek;
    public override bool CanWrite => true;
    public override long Length => _destination.CanSeek ? _destination.Length : _maximumPosition;

    public override long Position {
        get => _destination.CanSeek ? _destination.Position : _maximumPosition;
        set {
            if (!_destination.CanSeek) throw new NotSupportedException();
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
        long position = _destination.Seek(offset, origin);
        EnsureWithinLimit(position);
        return position;
    }

    public override void SetLength(long value) {
        EnsureWithinLimit(value);
        _destination.SetLength(value);
        if (_maximumPosition > value) _maximumPosition = value;
    }

    public override void Write(byte[] buffer, int offset, int count) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
        long end = checked(Position + count);
        EnsureWithinLimit(end);
        _destination.Write(buffer, offset, count);
        if (end > _maximumPosition) _maximumPosition = end;
    }

    public override void WriteByte(byte value) {
        long end = checked(Position + 1);
        EnsureWithinLimit(end);
        _destination.WriteByte(value);
        if (end > _maximumPosition) _maximumPosition = end;
    }

    public override async Task WriteAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
        long end = checked(Position + count);
        EnsureWithinLimit(end);
        await _destination.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
        if (end > _maximumPosition) _maximumPosition = end;
    }

#if NETSTANDARD2_1_OR_GREATER || NETCOREAPP
    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken = default) {
        long end = checked(Position + buffer.Length);
        EnsureWithinLimit(end);
        await _destination.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
        if (end > _maximumPosition) _maximumPosition = end;
    }
#endif

    protected override void Dispose(bool disposing) {
        // The caller owns the destination stream.
        base.Dispose(disposing);
    }

    private void EnsureWithinLimit(long length) {
        if (length < 0 || length > _maximumLength) {
            throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes), length, _maximumLength);
        }
    }
}
