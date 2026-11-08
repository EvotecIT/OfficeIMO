namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Defers the bounded worksheet compression handoff until a streamed sheet is large enough.</summary>
internal sealed class XlsbStreamingCompression(Stream destination) : Stream {
    private const int SynchronousPrefixLength = 1024 * 1024;
    private XlsbCompressionPipeline? _pipeline;
    private int _prefixBytes;
    private bool _disposed;
    private bool _complete;

    public override void Write(byte[] buffer, int offset, int count) {
        if (_disposed || _complete) throw new ObjectDisposedException(nameof(XlsbStreamingCompression));
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
        if (_pipeline == null && _prefixBytes < SynchronousPrefixLength) {
            destination.Write(buffer, offset, count);
            _prefixBytes += count;
            return;
        }
        _pipeline ??= new XlsbCompressionPipeline(destination);
        _pipeline.Write(buffer, offset, count);
    }

    /// <summary>Joins the compressor before the owning ZIP entry can be closed.</summary>
    internal void Complete() {
        if (_disposed) throw new ObjectDisposedException(nameof(XlsbStreamingCompression));
        if (_complete) return;
        _pipeline?.Complete();
        _complete = true;
    }

    protected override void Dispose(bool disposing) {
        if (_disposed) return;
        _disposed = true;
        if (disposing) _pipeline?.Dispose();
        base.Dispose(disposing);
    }

    public override void Flush() => _pipeline?.Flush();
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => !_disposed && !_complete;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
}
