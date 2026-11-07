namespace OfficeIMO.Email.Store;

/// <summary>Reads one indexed RFC message across linked DBX fragments without owning the database stream.</summary>
internal sealed class DbxMessageStream : Stream {
    private readonly DbxDatabase _database;
    private readonly EmailStoreReaderOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly byte[] _header = new byte[16];
    private readonly HashSet<uint> _seen = new HashSet<uint>();
    private uint _nextFragment;
    private long _payloadOffset;
    private int _remaining;
    private long _messageLength;

    internal DbxMessageStream(DbxDatabase database, uint firstFragment, EmailStoreReaderOptions options,
        CancellationToken cancellationToken) {
        if (firstFragment == 0) throw DbxDatabase.Invalid("The indexed message has no local payload.");
        _database = database;
        _nextFragment = firstFragment;
        _options = options;
        _cancellationToken = cancellationToken;
    }

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

    public override int Read(byte[] buffer, int offset, int count) {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
        if (count == 0) return 0;
        // Stop at the fragment boundary so callers reading a prefix need not open the next link.
        if (_remaining == 0) {
            if (_nextFragment == 0) return 0;
            OpenFragment();
        }
        int read = Math.Min(_remaining, count);
        _database.ReadAt(_payloadOffset, buffer, offset, read, _cancellationToken);
        _payloadOffset += read;
        _remaining -= read;
        return read;
    }

    public override void Flush() { }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    private void OpenFragment() {
        uint address = _nextFragment;
        if (_seen.Count >= _options.MaxNodeCount)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxNodeCount),
                _seen.Count + 1L, _options.MaxNodeCount);
        if (!_seen.Add(address)) throw DbxDatabase.Invalid("A message fragment chain contains a cycle.");

        _database.ReadObject(address, _header, _cancellationToken);
        uint allocated = DbxDatabase.UInt32(_header, 4);
        uint used = DbxDatabase.UInt32(_header, 8);
        if (used == 0 || used > allocated || used > int.MaxValue)
            throw DbxDatabase.Invalid("A message fragment declares an invalid payload length.");
        _payloadOffset = (long)address + _header.Length;
        _database.ValidateRange(_payloadOffset, allocated);
        _messageLength += used;
        if (_messageLength > _options.MaxMessageBytes)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxMessageBytes),
                _messageLength, _options.MaxMessageBytes);
        _remaining = (int)used;
        _nextFragment = DbxDatabase.UInt32(_header, 12);
    }
}
