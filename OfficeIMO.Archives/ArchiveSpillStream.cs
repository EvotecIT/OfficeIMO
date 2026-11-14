namespace OfficeIMO.Archives;

/// <summary>Seekable staging buffer that keeps small content in memory and moves to a private temporary file above 8 MiB.</summary>
internal sealed class ArchiveSpillStream : Stream {
    private const int MemoryLimit = 8 * 1024 * 1024;
    private Stream _inner = new MemoryStream();

    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => true;
    public override long Length => _inner.Length;
    public override long Position { get => _inner.Position; set => _inner.Position = value; }
    public override void Flush() => _inner.Flush();
    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    public override void SetLength(long value) => _inner.SetLength(value);

    public override void Write(byte[] buffer, int offset, int count) {
        if (_inner is MemoryStream memory && Math.Max(memory.Length, memory.Position + count) > MemoryLimit) {
            Stream file = ArchiveTemporaryFile.Create();
            try {
                long position = memory.Position;
                memory.WriteTo(file);
                file.Position = position;
            } catch { file.Dispose(); throw; }
            memory.Dispose(); _inner = file;
        }
        _inner.Write(buffer, offset, count);
    }

    internal void CopyTo(Stream destination, CancellationToken token) {
        if (_inner is MemoryStream memory) { memory.WriteTo(destination); return; }
        _inner.Position = 0;
        var buffer = new byte[81920];
        int read;
        while ((read = _inner.Read(buffer, 0, buffer.Length)) > 0) {
            token.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, read);
        }
    }

    protected override void Dispose(bool disposing) {
        if (disposing) _inner.Dispose();
        base.Dispose(disposing);
    }
}
