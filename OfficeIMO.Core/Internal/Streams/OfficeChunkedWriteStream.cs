namespace OfficeIMO.Core.Internal;

/// <summary>
/// Collects content of unknown length in growing chunks and returns it as one exact-length array.
/// </summary>
/// <remarks>
/// A growing <see cref="MemoryStream"/> reallocates and copies its buffer on every doubling and then copies once more
/// for <see cref="MemoryStream.ToArray"/>, allocating about three times the content. Chunks are never copied until
/// <see cref="ToArray"/>, so reading a non-seekable source allocates about twice its length.
/// </remarks>
internal sealed class OfficeChunkedWriteStream : Stream {
    private const int FirstChunkSize = 81920;
    private const int MaximumChunkSize = 4 * 1024 * 1024;

    private readonly List<byte[]> _chunks = new();
    private byte[]? _current;
    private int _currentCount;
    private long _length;

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => _length;

    public override long Position {
        get => _length;
        set => throw new NotSupportedException();
    }

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) {
        ArgumentNullException.ThrowIfNull(buffer);
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
        if (checked(_length + count) > int.MaxValue) {
            throw new InvalidDataException("Stream content exceeds the supported in-memory size.");
        }
        while (count > 0) {
            if (_current == null || _currentCount == _current.Length) {
                int size = _current == null ? FirstChunkSize : Math.Min(MaximumChunkSize, _current.Length * 2);
                _current = new byte[size];
                _currentCount = 0;
                _chunks.Add(_current);
            }
            int copy = Math.Min(count, _current.Length - _currentCount);
            Buffer.BlockCopy(buffer, offset, _current, _currentCount, copy);
            _currentCount += copy;
            _length += copy;
            offset += copy;
            count -= copy;
        }
    }

    public override void WriteByte(byte value) => Write(new[] { value }, 0, 1);

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        Write(buffer, offset, count);
        return Task.CompletedTask;
    }

    /// <summary>Returns the collected content as one array and releases the chunks.</summary>
    public byte[] ToArray() {
        if (_length == 0) return Array.Empty<byte>();
        var result = new byte[_length];
        int written = 0;
        foreach (byte[] chunk in _chunks) {
            int count = Math.Min(chunk.Length, result.Length - written);
            Buffer.BlockCopy(chunk, 0, result, written, count);
            written += count;
        }
        _chunks.Clear();
        _current = null;
        _currentCount = 0;
        _length = 0;
        return result;
    }
}
