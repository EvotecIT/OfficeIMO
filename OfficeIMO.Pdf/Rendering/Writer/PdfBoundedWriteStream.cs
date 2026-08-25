namespace OfficeIMO.Pdf;

/// <summary>Forwards writes while rejecting any byte that would cross an owned-output ceiling.</summary>
internal sealed class PdfBoundedWriteStream : Stream {
    private readonly Stream _inner;
    private readonly long? _maximumBytes;
    private readonly string _limitMessage;

    internal PdfBoundedWriteStream(Stream inner, long? maximumBytes, string limitMessage) {
        Guard.NotNull(inner, nameof(inner));
        Guard.NotNull(limitMessage, nameof(limitMessage));
        _inner = inner;
        _maximumBytes = maximumBytes;
        _limitMessage = limitMessage;
    }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => _inner.Length;
    public override long Position { get => _inner.Position; set => throw new NotSupportedException(); }
    public override void Flush() => _inner.Flush();
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) {
        EnsureWithinLimit(count);
        _inner.Write(buffer, offset, count);
    }

    public override void WriteByte(byte value) {
        EnsureWithinLimit(1L);
        _inner.WriteByte(value);
    }

    private void EnsureWithinLimit(long addedBytes) {
        if (_maximumBytes.HasValue && _inner.Position > _maximumBytes.Value - addedBytes) {
            throw new InvalidDataException(_limitMessage);
        }
    }

    protected override void Dispose(bool disposing) {
        if (disposing) _inner.Flush();
        base.Dispose(disposing);
    }
}
