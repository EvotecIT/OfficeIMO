namespace OfficeIMO.Reader;

/// <summary>Replays an already-read prefix while preserving ownership of the underlying stream.</summary>
internal sealed class ReaderPrefixStream : Stream {
    private readonly Stream _source;
    private readonly byte[] _prefix;
    private int _position;
    private readonly int _end;
    internal ReaderPrefixStream(Stream source, byte[] prefix, int start, int end) {
        _source = source; _prefix = prefix; _position = start; _end = end;
    }
    public override int Read(byte[] buffer, int offset, int count) {
        int take = Math.Min(count, _end - _position);
        if (take > 0) { Array.Copy(_prefix, _position, buffer, offset, take); _position += take; return take; }
        return _source.Read(buffer, offset, count);
    }
    public override bool CanRead => _source.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override void Flush() { }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    // Disposing the decoder never closes the caller's stream.
}
