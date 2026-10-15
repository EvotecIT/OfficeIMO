namespace OfficeIMO.OpenDocument;

/// <summary>One read from the native owner's immutable compressed snapshot; owns only its transient archive and entry stream.</summary>
internal sealed class OdfDeferredEntryStream : Stream {
    private readonly MemoryStream _input;
    private readonly ZipArchive _archive;
    private readonly Stream _entry;
    private readonly CancellationToken _token;
    private readonly long _expectedLength;
    private readonly uint _expectedChecksum;
    private uint _checksum = OfficeIMO.Drawing.OfficePngCrc32.Begin();
    private long _observedLength;
    private bool _complete, _disposed;
    private OdfDeferredEntryStream(MemoryStream input, ZipArchive archive, Stream entry, long length, uint checksum, CancellationToken token) {
        _input = input; _archive = archive; _entry = entry; _expectedLength = length; _expectedChecksum = checksum; _token = token;
    }
    internal static Stream Open(byte[] snapshot, int ordinal, string path, long length,
        OfficeIMO.Provenance.OfficeProvenanceZip.OfficeProvenanceZipEntryMetadata metadata, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var input = new MemoryStream(snapshot, writable: false);
        ZipArchive? archive = null;
        try {
            archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
            ZipArchiveEntry entry = archive.Entries[ordinal];
            if (entry.FullName != path || entry.Length != length)
                throw new InvalidDataException("Deferred ODF source entry differs from its validated directory.");
            token.ThrowIfCancellationRequested();
            Stream payload = OfficeIMO.Core.Internal.OfficeArchiveSafety.OpenEntryPayload(snapshot, entry, metadata);
            return new OdfDeferredEntryStream(input, archive, payload, length, metadata.Checksum, token);
        } catch { archive?.Dispose(); input.Dispose(); throw; }
    }
    public override bool CanRead => !_disposed && _entry.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override int Read(byte[] buffer, int offset, int count) {
        if (_disposed) throw new ObjectDisposedException(nameof(OdfDeferredEntryStream));
        _token.ThrowIfCancellationRequested();
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
        if (count == 0 || _complete) return 0;
        int read = _entry.Read(buffer, offset, count);
        if (read > 0) {
            _observedLength = checked(_observedLength + read);
            _checksum = OfficeIMO.Drawing.OfficePngCrc32.Append(_checksum, buffer, offset, read);
            return read;
        }
        if (_observedLength != _expectedLength) throw new InvalidDataException("The decoded ODF payload differs from its declared length.");
        if (OfficeIMO.Drawing.OfficePngCrc32.Complete(_checksum) != _expectedChecksum)
            throw new InvalidDataException("The decoded ODF payload does not match its ZIP checksum.");
        _complete = true;
        return 0;
    }
    public override void Flush() => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) {
        if (disposing && !_disposed) { _disposed = true; _entry.Dispose(); _archive.Dispose(); _input.Dispose(); }
        base.Dispose(disposing);
    }
}
