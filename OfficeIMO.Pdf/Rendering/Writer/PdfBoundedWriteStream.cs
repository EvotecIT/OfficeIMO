namespace OfficeIMO.Pdf;

/// <summary>Forwards writes while rejecting any byte that would cross an owned-output ceiling.</summary>
internal sealed class PdfBoundedWriteStream : Stream {
    private readonly Stream _inner;
    private readonly long? _maximumBytes;
    private readonly string _limitMessage;
    private long _writtenBytes;

    internal PdfBoundedWriteStream(Stream inner, long? maximumBytes, string limitMessage) {
        Guard.NotNull(inner, nameof(inner));
        Guard.NotNull(limitMessage, nameof(limitMessage));
        _inner = inner;
        _maximumBytes = maximumBytes;
        _limitMessage = limitMessage;
        _writtenBytes = inner.CanSeek ? inner.Position : 0L;
    }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => _inner.CanSeek ? _inner.Length : _writtenBytes;
    public override long Position { get => _writtenBytes; set => throw new NotSupportedException(); }
    public override void Flush() => _inner.Flush();
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) {
        EnsureWithinLimit(count);
        _inner.Write(buffer, offset, count);
        _writtenBytes += count;
    }

    public override void WriteByte(byte value) {
        EnsureWithinLimit(1L);
        _inner.WriteByte(value);
        _writtenBytes++;
    }

    private void EnsureWithinLimit(long addedBytes) {
        if (_maximumBytes.HasValue && _writtenBytes > _maximumBytes.Value - addedBytes) {
            throw PdfOutputLimitErrors.Create(_limitMessage);
        }
    }

    protected override void Dispose(bool disposing) {
        if (disposing) _inner.Flush();
        base.Dispose(disposing);
    }
}
