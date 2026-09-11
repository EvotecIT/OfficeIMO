namespace OfficeIMO.Invoicing;

/// <summary>Enforces the invoice XML byte ceiling before the underlying buffer can grow beyond it.</summary>
internal sealed class InvoiceXmlOutputStream : Stream {
    private readonly MemoryStream _output = new MemoryStream();
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => _output.Length;
    public override long Position { get => _output.Position; set => throw new NotSupportedException(); }
    public override void Flush() => _output.Flush();
    public override void Write(byte[] buffer, int offset, int count) {
        if (count > InvoiceProfileDeclaration.MaximumXmlBytes - _output.Length)
            throw new InvalidDataException("Serialized invoice exceeds 16 MiB.");
        _output.Write(buffer, offset, count);
    }
    internal byte[] ToArray() => _output.ToArray();
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) { if (disposing) _output.Dispose(); base.Dispose(disposing); }
}
