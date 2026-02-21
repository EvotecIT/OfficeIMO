namespace OfficeIMO.Tests;

internal sealed class NonSeekableReadStream : Stream {
    private readonly Stream _inner;

    public NonSeekableReadStream(byte[] bytes) {
        _inner = new MemoryStream(bytes, writable: false);
    }

    public override bool CanRead => _inner.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() => _inner.Flush();
    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) {
        if (disposing) {
            _inner.Dispose();
        }

        base.Dispose(disposing);
    }
}
