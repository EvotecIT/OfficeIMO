namespace OfficeIMO.Email.Store;

/// <summary>Exposes one bounded seekable segment without owning the underlying source.</summary>
internal sealed class EmailStoreSegmentStream : Stream {
    private readonly Stream _source;
    private readonly long _start;
    private readonly long _length;
    private long _position;
    internal EmailStoreSegmentStream(Stream source, long start, long length) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead || !source.CanSeek) throw new ArgumentException("A readable seekable source is required.", nameof(source));
        if (start < 0 || length < 0 || start > source.Length - length) throw new ArgumentOutOfRangeException(nameof(length));
        _source = source;
        _start = start;
        _length = length;
    }
    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => false;
    public override long Length => _length;
    public override long Position {
        get => _position;
        set {
            if (value < 0 || value > _length) throw new ArgumentOutOfRangeException(nameof(value));
            _position = value;
        }
    }
    public override int Read(byte[] buffer, int offset, int count) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
        int bounded = (int)Math.Min(count, _length - _position);
        if (bounded == 0) return 0;
        long absolutePosition = _start + _position;
        if (_source.Position != absolutePosition) _source.Position = absolutePosition;
        int read = _source.Read(buffer, offset, bounded);
        _position += read;
        return read;
    }
    public override int ReadByte() {
        if (_position >= _length) return -1;
        long absolutePosition = _start + _position;
        if (_source.Position != absolutePosition) _source.Position = absolutePosition;
        int value = _source.ReadByte();
        if (value >= 0) _position++;
        return value;
    }
    public override long Seek(long offset, SeekOrigin origin) {
        Position = origin == SeekOrigin.Begin ? offset : origin == SeekOrigin.Current ? checked(_position + offset) :
            origin == SeekOrigin.End ? checked(_length + offset) : throw new ArgumentOutOfRangeException(nameof(origin));
        return _position;
    }
    public override void Flush() { }
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}
