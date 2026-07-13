namespace OfficeIMO.Tests;

/// <summary>Caller-owned readable/writable stream that deliberately cannot replace existing content.</summary>
internal sealed class NonSeekableReadWriteBuffer : Stream {
    private readonly MemoryStream _inner;

    public NonSeekableReadWriteBuffer(byte[] bytes) {
        _inner = new MemoryStream();
        _inner.Write(bytes, 0, bytes.Length);
        _inner.Position = 0;
    }

    public byte[] ToArray() => _inner.ToArray();

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => throw new NotSupportedException();
    public override long Position {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() => _inner.Flush();
    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => _inner.SetLength(value);
    public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

    protected override void Dispose(bool disposing) {
        if (disposing) {
            _inner.Dispose();
        }
        base.Dispose(disposing);
    }
}
