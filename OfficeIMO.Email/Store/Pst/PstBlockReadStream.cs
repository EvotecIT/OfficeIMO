namespace OfficeIMO.Email.Store;

/// <summary>Forward-only stream over decoded NDB leaf blocks, retaining at most one block at a time.</summary>
internal sealed class PstBlockReadStream : Stream {
    private readonly IEnumerator<byte[]> _blocks;
    private readonly EmailStoreSessionLifetime _lifetime;
    private readonly Action<int>? _bytesRead;
    private byte[]? _current;
    private int _currentOffset;
    private bool _disposed;

    internal PstBlockReadStream(Func<IEnumerable<byte[]>> blocks,
        EmailStoreSessionLifetime lifetime, Action<int>? bytesRead = null) {
        if (blocks == null) throw new ArgumentNullException(nameof(blocks));
        _lifetime = lifetime ?? throw new ArgumentNullException(nameof(lifetime));
        _bytesRead = bytesRead;
        _blocks = blocks().GetEnumerator();
    }

    public override bool CanRead => !_disposed;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) {
            throw new ArgumentOutOfRangeException(offset < 0 ? nameof(offset) : nameof(count));
        }
        ThrowIfDisposed();
        _lifetime.ThrowIfDisposed();
        if (count == 0) return 0;

        int written = 0;
        while (written < count) {
            if (_current == null || _currentOffset >= _current.Length) {
                if (!MoveNextNonEmpty()) break;
            }
            int available = _current!.Length - _currentOffset;
            int copy = Math.Min(available, count - written);
            Buffer.BlockCopy(_current, _currentOffset, buffer, offset + written, copy);
            _currentOffset += copy;
            written += copy;
        }
        if (written > 0) _bytesRead?.Invoke(written);
        return written;
    }

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(Read(buffer, offset, count));
    }

    public override void Flush() { }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) {
        if (!_disposed) {
            _disposed = true;
            if (disposing) _blocks.Dispose();
            _current = null;
        }
        base.Dispose(disposing);
    }

    private bool MoveNextNonEmpty() {
        while (_blocks.MoveNext()) {
            _current = _blocks.Current ?? throw new InvalidDataException("A PST data block was null.");
            _currentOffset = 0;
            if (_current.Length > 0) return true;
        }
        _current = null;
        return false;
    }

    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(PstBlockReadStream));
    }
}
