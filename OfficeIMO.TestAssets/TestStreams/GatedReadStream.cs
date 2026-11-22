namespace OfficeIMO.Tests;

/// <summary>
/// A seekable stream whose data is available only through asynchronous reads once the test opens the gate.
/// </summary>
internal sealed class GatedReadStream : Stream {
    private readonly MemoryStream _inner;
    private readonly Task _gate;
    private readonly int _synchronousPrefixLength;

    internal GatedReadStream(byte[] bytes, Task gate, int synchronousPrefixLength = 0) {
        _inner = new MemoryStream(bytes, writable: false);
        _gate = gate;
        _synchronousPrefixLength = synchronousPrefixLength;
    }

    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => false;
    public override long Length => _inner.Length;
    public override long Position { get => _inner.Position; set => _inner.Position = value; }

    public override int Read(byte[] buffer, int offset, int count) {
        if (_inner.Position + count > _synchronousPrefixLength) throw new NotSupportedException("Use ReadAsync for package content.");
        return _inner.Read(buffer, offset, count);
    }

    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
        await _gate.ConfigureAwait(false);
        return await _inner.ReadAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
    }

    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    public override void Flush() { }
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) {
        if (disposing) _inner.Dispose();
        base.Dispose(disposing);
    }
}
