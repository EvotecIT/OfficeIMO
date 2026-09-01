namespace OfficeIMO.Workflows;

/// <summary>Counts a seekable serialization without retaining its bytes and fails at a configured boundary.</summary>
internal sealed class OfficeWorkflowBoundedCountingStream : Stream {
    private readonly long _maximumBytes;
    private long _length;
    private long _position;

    internal OfficeWorkflowBoundedCountingStream(long maximumBytes) {
        if (maximumBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
    }

    public override bool CanRead => false;
    public override bool CanSeek => true;
    public override bool CanWrite => true;
    public override long Length => _length;
    public override long Position {
        get => _position;
        set {
            EnsureWithinLimit(value);
            _position = value;
        }
    }

    public override void Flush() { }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) {
        long basis = origin switch {
            SeekOrigin.Begin => 0L,
            SeekOrigin.Current => _position,
            SeekOrigin.End => _length,
            _ => throw new ArgumentOutOfRangeException(nameof(origin))
        };
        long target = checked(basis + offset);
        if (target < 0L) throw new IOException("Cannot seek before the beginning of the stream.");
        EnsureWithinLimit(target);
        _position = target;
        return target;
    }

    public override void SetLength(long value) {
        EnsureWithinLimit(value);
        _length = value;
        if (_position > value) _position = value;
    }

    public override void Write(byte[] buffer, int offset, int count) {
        ArgumentNullException.ThrowIfNull(buffer);
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
        Advance(count);
    }

    public override void Write(ReadOnlySpan<byte> buffer) => Advance(buffer.Length);

    public override void WriteByte(byte value) => Advance(1);

    private void Advance(int count) {
        long end;
        try {
            end = checked(_position + count);
        } catch (OverflowException) {
            throw CreateLimitException();
        }
        EnsureWithinLimit(end);
        _position = end;
        if (end > _length) _length = end;
    }

    private void EnsureWithinLimit(long value) {
        if (value < 0L) throw new ArgumentOutOfRangeException(nameof(value));
        if (value > _maximumBytes) throw CreateLimitException();
    }

    private InvalidOperationException CreateLimitException() => new(
        $"Generated artifact exceeded the configured {_maximumBytes:N0}-byte output limit while it was being serialized.");
}

/// <summary>Seekable write-through stream that rejects growth beyond a configured artifact boundary.</summary>
internal sealed class OfficeWorkflowBoundedWriteStream : Stream {
    private readonly Stream _inner;
    private readonly long _maximumBytes;
    private readonly bool _leaveOpen;

    internal OfficeWorkflowBoundedWriteStream(Stream inner, long maximumBytes, bool leaveOpen = false) {
        ArgumentNullException.ThrowIfNull(inner);
        if (!inner.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(inner));
        if (maximumBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _inner = inner;
        _maximumBytes = maximumBytes;
        _leaveOpen = leaveOpen;
    }

    public override bool CanRead => false;
    public override bool CanSeek => _inner.CanSeek;
    public override bool CanWrite => true;
    public override long Length => _inner.Length;
    public override long Position {
        get => _inner.Position;
        set {
            EnsureWithinLimit(value);
            _inner.Position = value;
        }
    }

    public override void Flush() => _inner.Flush();
    public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) {
        long previous = _inner.Position;
        long target = _inner.Seek(offset, origin);
        try {
            EnsureWithinLimit(target);
            return target;
        } catch {
            _inner.Position = previous;
            throw;
        }
    }

    public override void SetLength(long value) {
        EnsureWithinLimit(value);
        _inner.SetLength(value);
    }

    public override void Write(byte[] buffer, int offset, int count) {
        EnsureWriteWithinLimit(count);
        _inner.Write(buffer, offset, count);
    }

    public override void Write(ReadOnlySpan<byte> buffer) {
        EnsureWriteWithinLimit(buffer.Length);
        _inner.Write(buffer);
    }

    public override void WriteByte(byte value) {
        EnsureWriteWithinLimit(1);
        _inner.WriteByte(value);
    }

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
        EnsureWriteWithinLimit(count);
        return _inner.WriteAsync(buffer, offset, count, cancellationToken);
    }

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) {
        EnsureWriteWithinLimit(buffer.Length);
        return _inner.WriteAsync(buffer, cancellationToken);
    }

    protected override void Dispose(bool disposing) {
        if (disposing && !_leaveOpen) _inner.Dispose();
        base.Dispose(disposing);
    }

    public override async ValueTask DisposeAsync() {
        if (!_leaveOpen) await _inner.DisposeAsync().ConfigureAwait(false);
        GC.SuppressFinalize(this);
    }

    private void EnsureWriteWithinLimit(int count) {
        long end;
        try {
            end = checked(Position + count);
        } catch (OverflowException) {
            throw CreateLimitException();
        }
        EnsureWithinLimit(Math.Max(Length, end));
    }

    private void EnsureWithinLimit(long value) {
        if (value < 0L) throw new ArgumentOutOfRangeException(nameof(value));
        if (value > _maximumBytes) throw CreateLimitException();
    }

    private InvalidOperationException CreateLimitException() => new(
        $"Generated artifact exceeded the configured {_maximumBytes:N0}-byte output limit while it was being serialized.");
}
