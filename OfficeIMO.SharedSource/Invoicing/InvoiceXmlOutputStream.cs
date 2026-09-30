using System;
using System.IO;

namespace OfficeIMO.Internal.Invoicing;

/// <summary>Bounds serialized XML before an escaping-heavy edit can allocate beyond the input contract.</summary>
internal sealed class InvoiceXmlOutputStream : Stream {
    private readonly MemoryStream _buffer = new MemoryStream();
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => _buffer.CanWrite;
    public override long Length => _buffer.Length;
    public override long Position { get => _buffer.Position; set => throw new NotSupportedException(); }
    public override void Flush() => _buffer.Flush();
    public override void Write(byte[] buffer, int offset, int count) {
        if (_buffer.Length + count > InvoiceProfileDeclaration.MaximumXmlBytes) throw new InvalidDataException("Edited invoice XML exceeds the 16 MiB output limit.");
        _buffer.Write(buffer, offset, count);
    }
    internal byte[] ToArray() => _buffer.ToArray();
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) { if (disposing) _buffer.Dispose(); base.Dispose(disposing); }
}
