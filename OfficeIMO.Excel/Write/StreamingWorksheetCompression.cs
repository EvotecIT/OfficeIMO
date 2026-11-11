
namespace OfficeIMO.Excel;

/// <summary>Defers the bounded worksheet compression handoff until a streamed sheet is large enough.</summary>
/// <remarks>With an <paramref name="owner"/> archive the sheet borrows that package's compressor, so a workbook of many
/// sheets starts one worker instead of one per sheet.</remarks>
internal sealed class StreamingWorksheetCompression(Stream destination, ExcelPackageArchive? owner = null) : Stream {
    // Smaller sheets stay on the caller. The package keeps its compressor for later sheets, so only the first large
    // sheet pays to start the worker.
    private const int SynchronousPrefixLength = 64 * 1024;
    private WorksheetCompressionPipeline? _pipeline;
    private int _prefixBytes;
    private bool _disposed;
    private bool _complete;

    public override void Write(byte[] buffer, int offset, int count) {
        if (_disposed || _complete) throw new ObjectDisposedException(nameof(StreamingWorksheetCompression));
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
        if (_pipeline == null && _prefixBytes < SynchronousPrefixLength) {
            destination.Write(buffer, offset, count);
            _prefixBytes += count;
            return;
        }
        _pipeline ??= owner?.AcquireCompression(destination) ?? new WorksheetCompressionPipeline(destination);
        _pipeline.Write(buffer, offset, count);
    }

    /// <summary>Joins the compressor before the owning ZIP entry can be closed.</summary>
    internal void Complete() {
        if (_disposed) throw new ObjectDisposedException(nameof(StreamingWorksheetCompression));
        if (_complete) return;
        if (owner != null) _pipeline?.Drain();
        else _pipeline?.Complete();
        _complete = true;
    }

    protected override void Dispose(bool disposing) {
        if (_disposed) return;
        _disposed = true;
        if (disposing && _pipeline != null) {
            // A borrowed compressor stays with the package only after this entry drained completely.
            if (owner == null) _pipeline.Dispose();
            else if (!_complete) owner.AbandonCompression();
        }
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
