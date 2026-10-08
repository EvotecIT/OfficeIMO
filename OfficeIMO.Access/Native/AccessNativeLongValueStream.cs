using OfficeIMO.Drawing;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access;

/// <summary>Incremental read over one bounded long-value chain. No whole-payload allocation occurs.</summary>
internal sealed class AccessNativeLongValueStream : Stream {
    private readonly AccessNativeDatabase _database;
    private readonly CancellationToken _cancellation;
    private readonly int _length;
    private readonly uint _kind;
    private uint _next;
    private OfficeByteView _chunk;
    private int _chunkOffset, _depth, _position;
    private bool _disposed;
    private readonly HashSet<uint> _visited = new HashSet<uint>();
    internal AccessNativeLongValueStream(AccessNativeDatabase database, OfficeByteView descriptor, CancellationToken cancellation, int? maximumBytes = null) {
        _database = database; _cancellation = cancellation;
        if (descriptor.Length == 0) { _length = 0; _kind = 2; return; }
        uint header = U32(descriptor, 0); _length = checked((int)(header & 0x3fffffff)); _kind = header >> 30;
        if (_length > (maximumBytes ?? database.MaxValueBytes)) throw new InvalidDataException("Native Access long value exceeds its value limit.");
        if (_kind == 2) _chunk = Slice(descriptor, 12, _length);
        else {
            if (_kind > 1 || descriptor.Length != 12) throw new InvalidDataException("Native Access long-value descriptor is invalid.");
            _next = U32(descriptor, 4);
            if (_length == 0 && _next != 0) throw new InvalidDataException("Native Access empty long value has a dangling chain.");
        }
    }
    private void Check() { if (_disposed) throw new ObjectDisposedException(nameof(AccessNativeLongValueStream)); _database.Document.EnsureNotDisposed(); _cancellation.ThrowIfCancellationRequested(); }
    public override bool CanRead => !_disposed;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length { get { Check(); return _length; } }
    public override long Position { get { Check(); return _position; } set => throw new NotSupportedException(); }
    public override int Read(byte[] buffer, int offset, int count) => ReadCore(buffer, offset, count, default);
    private int ReadCore(byte[] buffer, int offset, int count, CancellationToken cancellation) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
        Check(); cancellation.ThrowIfCancellationRequested(); int total = 0;
        while (count > 0 && _position < _length) {
            Check(); cancellation.ThrowIfCancellationRequested();
            if (_chunkOffset == _chunk.Length) {
                if (_next == 0 || _depth++ >= _database.MaxChainLength || !_visited.Add(_next)) throw new InvalidDataException("Native Access long-value chain is truncated, cyclic or exceeds its limit.");
                int page = checked((int)(_next >> 8)); var pageBytes = _database.Page(page, 1);
                if (pageBytes[4] != 'L' || pageBytes[5] != 'V' || pageBytes[6] != 'A' || pageBytes[7] != 'L') throw new InvalidDataException("Native Access long-value chain refers to a non-LVAL data page.");
                var row = _database.Row(page, (int)(_next & 255), false, _cancellation);
                int prefix = _kind == 0 ? 4 : 0;
                if (row.Length <= prefix) throw new InvalidDataException("Native Access long-value chain makes no progress.");
                _chunk = Slice(row, prefix, row.Length - prefix); _chunkOffset = 0; _next = _kind == 0 ? U32(row, 0) : 0;
            }
            int amount = Math.Min(count, Math.Min(_chunk.Length - _chunkOffset, _length - _position));
            for (int i = 0; i < amount; i++) buffer[offset + i] = _chunk[_chunkOffset + i];
            _chunkOffset += amount; _position += amount; total += amount; offset += amount; count -= amount;
        }
        if (_position == _length && _next != 0) throw new InvalidDataException("Native Access long-value chain exceeds its declared length.");
        return total;
    }
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
        return Task.FromResult(ReadCore(buffer, offset, count, cancellationToken));
    }
    public override void Flush() { Check(); }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) { _disposed = true; _chunk = default; _visited.Clear(); base.Dispose(disposing); }
}
