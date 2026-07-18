namespace OfficeIMO.Pdf;

/// <summary>Hashes sequential PDF output while preserving the destination stream contract.</summary>
internal sealed class PdfPipelineHashingStream : Stream {
    private readonly Stream _destination;
    private readonly System.Security.Cryptography.HashAlgorithm _sha256;
    private long _bytesWritten;
    private bool _completed;

    internal PdfPipelineHashingStream(Stream destination) {
        Guard.NotNull(destination, nameof(destination));
        _destination = destination;
        _sha256 = System.Security.Cryptography.SHA256.Create();
    }

    public override bool CanRead => false;
    public override bool CanSeek => _destination.CanSeek;
    public override bool CanWrite => _destination.CanWrite;
    public override long Length => _destination.Length;
    public override long Position {
        get => _destination.Position;
        set => _destination.Position = value;
    }

    internal PdfArtifactSnapshot Complete(int? pageCount) {
        if (!_completed) {
            _sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            _completed = true;
        }

        return PdfArtifactSnapshot.FromDigest(
            _bytesWritten,
            ToLowerHex(_sha256.Hash ?? Array.Empty<byte>()),
            pageCount);
    }

    public override void Flush() => _destination.Flush();
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => _destination.Seek(offset, origin);
    public override void SetLength(long value) => _destination.SetLength(value);

    public override void Write(byte[] buffer, int offset, int count) {
        if (_completed) {
            throw new InvalidOperationException("PDF output hash has already been completed.");
        }

        _destination.Write(buffer, offset, count);
        _sha256.TransformBlock(buffer, offset, count, null, 0);
        _bytesWritten += count;
    }

    protected override void Dispose(bool disposing) {
        if (disposing) {
            _sha256.Dispose();
        }

        base.Dispose(disposing);
    }

    private static string ToLowerHex(byte[] bytes) {
        const string hex = "0123456789abcdef";
        char[] chars = new char[bytes.Length * 2];
        for (int i = 0; i < bytes.Length; i++) {
            chars[i * 2] = hex[bytes[i] >> 4];
            chars[(i * 2) + 1] = hex[bytes[i] & 0x0F];
        }

        return new string(chars);
    }
}
