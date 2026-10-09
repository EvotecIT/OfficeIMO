using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal;

/// <summary>Owns a read-only stream whose operations remain subject to its document lifetime and originating cancellation.</summary>
internal sealed class OfficeDocumentReadStream : Stream {
    private readonly Stream _inner;
    private readonly Action _ensureDocumentAlive;
    private readonly CancellationToken _cancellation;
    private bool _disposed;

    internal OfficeDocumentReadStream(Stream inner, Action ensureDocumentAlive, CancellationToken cancellation) {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _ensureDocumentAlive = ensureDocumentAlive ?? throw new ArgumentNullException(nameof(ensureDocumentAlive));
        _cancellation = cancellation;
    }
    private void Check() {
        if (_disposed) throw new ObjectDisposedException(nameof(OfficeDocumentReadStream));
        _ensureDocumentAlive(); _cancellation.ThrowIfCancellationRequested();
    }
    public override bool CanRead => !_disposed && _inner.CanRead;
    public override bool CanSeek => !_disposed && _inner.CanSeek;
    public override bool CanWrite => false;
    public override long Length { get { Check(); return _inner.Length; } }
    public override long Position { get { Check(); return _inner.Position; } set { Check(); _inner.Position = value; } }
    public override int Read(byte[] buffer, int offset, int count) { Check(); return _inner.Read(buffer, offset, count); }
    public override int ReadByte() { Check(); return _inner.ReadByte(); }
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
        Check(); cancellationToken.ThrowIfCancellationRequested(); return Task.FromResult(_inner.Read(buffer, offset, count));
    }
    public override long Seek(long offset, SeekOrigin origin) { Check(); return _inner.Seek(offset, origin); }
    public override void Flush() { Check(); }
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) {
        if (!_disposed) { _disposed = true; if (disposing) _inner.Dispose(); }
        base.Dispose(disposing);
    }
}
