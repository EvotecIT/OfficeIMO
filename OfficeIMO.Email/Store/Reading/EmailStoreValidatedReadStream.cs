namespace OfficeIMO.Email.Store;

/// <summary>Owns a recovered part stream and verifies its pinned bytes before its read scope ends.</summary>
internal sealed class EmailStoreValidatedReadStream : Stream {
    private readonly Stream _input;
    private readonly EmailStoreSourceGuard _guard;
    private readonly CancellationToken _cancellationToken;
    private bool _disposed;

    internal EmailStoreValidatedReadStream(Stream input, EmailStoreSourceGuard guard, CancellationToken cancellationToken) {
        _input = input;
        _guard = guard;
        _cancellationToken = cancellationToken;
    }

    public override bool CanRead => !_disposed && _input.CanRead;
    public override bool CanSeek => !_disposed && _input.CanSeek;
    public override bool CanWrite => false;
    public override long Length => _input.Length;
    public override long Position { get => _input.Position; set => _input.Position = value; }
    public override int Read(byte[] buffer, int offset, int count) => _input.Read(buffer, offset, count);
    public override int ReadByte() => _input.ReadByte();
    public override long Seek(long offset, SeekOrigin origin) => _input.Seek(offset, origin);
    public override void Flush() { }
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) {
        if (disposing && !_disposed) {
            _disposed = true;
            try {
                // A cancelled projection cannot escape; avoid another full scan during cancellation cleanup.
                if (!_cancellationToken.IsCancellationRequested) _guard.Validate(_input, _cancellationToken);
            } finally { _input.Dispose(); }
        }
        base.Dispose(disposing);
    }
}
