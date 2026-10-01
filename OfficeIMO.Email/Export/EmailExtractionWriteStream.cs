using System.Security.Cryptography;

namespace OfficeIMO.Email;

/// <summary>Observes serialized bytes and cancellation without taking ownership of the destination.</summary>
internal sealed class EmailExtractionWriteStream : Stream {
    private readonly Stream _output;
    private readonly CancellationToken _cancellation;
    private readonly Action<int> _written;
    private readonly IncrementalHash _hash;

    internal EmailExtractionWriteStream(Stream output, CancellationToken cancellation, Action<int> written, IncrementalHash hash) {
        _output = output; _cancellation = cancellation; _written = written; _hash = hash;
    }
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override void Flush() { _cancellation.ThrowIfCancellationRequested(); _output.Flush(); }
    public override void Write(byte[] buffer, int offset, int count) {
        _cancellation.ThrowIfCancellationRequested();
        _output.Write(buffer, offset, count);
        _written(count);
        _hash.AppendData(buffer, offset, count);
    }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
}
