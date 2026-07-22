namespace OfficeIMO.Email.Store;

/// <summary>Counts bytes actually produced by OLM ZIP decompression.</summary>
internal sealed class OlmDecodedArchiveBudget {
    private readonly long _maximumBytes;
    private readonly object _gate = new object();
    private long _decodedBytes;

    internal OlmDecodedArchiveBudget(long maximumBytes) {
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
    }

    internal void Add(long bytes) {
        if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
        if (bytes == 0) return;
        lock (_gate) {
            if (bytes > _maximumBytes - _decodedBytes) {
                long actual = bytes > long.MaxValue - _decodedBytes
                    ? long.MaxValue
                    : _decodedBytes + bytes;
                throw new EmailStoreLimitExceededException(
                    nameof(EmailStoreReaderOptions.MaxArchiveDecodedBytes),
                    actual, _maximumBytes);
            }
            _decodedBytes += bytes;
        }
    }
}

/// <summary>Bounds one decoded ZIP entry while charging a read-wide aggregate budget.</summary>
internal sealed class OlmDecodedEntryStream : Stream {
    private readonly Stream _inner;
    private readonly long _maximumBytes;
    private readonly string _limitName;
    private readonly OlmDecodedArchiveBudget _aggregateBudget;
    private long _decodedBytes;

    internal OlmDecodedEntryStream(Stream inner, long maximumBytes,
        string limitName, OlmDecodedArchiveBudget aggregateBudget) {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
        _limitName = limitName ?? throw new ArgumentNullException(nameof(limitName));
        _aggregateBudget = aggregateBudget ?? throw new ArgumentNullException(nameof(aggregateBudget));
    }

    public override bool CanRead => _inner.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position {
        get => _decodedBytes;
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) =>
        Count(_inner.Read(buffer, offset, count));

    public override int ReadByte() {
        int value = _inner.ReadByte();
        if (value >= 0) Count(1);
        return value;
    }

    protected override void Dispose(bool disposing) {
        if (disposing) _inner.Dispose();
        base.Dispose(disposing);
    }

    private int Count(int bytes) {
        if (bytes <= 0) return bytes;
        if (bytes > _maximumBytes - _decodedBytes) {
            long actual = bytes > long.MaxValue - _decodedBytes
                ? long.MaxValue
                : _decodedBytes + bytes;
            throw new EmailStoreLimitExceededException(
                _limitName,
                actual, _maximumBytes);
        }
        _decodedBytes += bytes;
        _aggregateBudget.Add(bytes);
        return bytes;
    }

    public override void Flush() => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}
