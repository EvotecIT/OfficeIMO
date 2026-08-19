namespace OfficeIMO.Email.Store;

/// <summary>Append-only fixed-record block index kept off the managed heap.</summary>
internal sealed class PstWriterBlockJournal : IDisposable {
    private const int RecordLength = 24;
    private const string ReferenceCountMigrationSuffix = ".refcounts";
    private readonly string _path;
    private readonly string _referenceCountMigrationPath;
    private readonly FileStream _stream;
    private readonly BinaryWriter _writer;
    private readonly byte[] _referenceCountBuffer = new byte[sizeof(int)];
    private bool _deleteOnDispose = true;
    private bool _disposed;

    internal PstWriterBlockJournal(string path, bool resume = false, long recordCount = 0) {
        _path = path;
        _referenceCountMigrationPath = string.Concat(path, ReferenceCountMigrationSuffix);
        _stream = new FileStream(path, resume ? FileMode.Open : FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.Read, 64 * 1024, FileOptions.SequentialScan);
        if (resume) {
            long length = checked(recordCount * RecordLength);
            if (_stream.Length < length) throw new InvalidDataException("The PST block journal is truncated.");
            _stream.SetLength(length);
        }
        _writer = new BinaryWriter(_stream, Encoding.UTF8, leaveOpen: true);
    }

    internal long Count => _stream.Length / RecordLength;

    internal void Add(PstWriterBlock block) {
        _stream.Position = _stream.Length;
        _writer.Write(block.Bid);
        _writer.Write(block.Offset);
        _writer.Write(block.Length);
        _writer.Write(block.ReferenceCount);
    }

    internal void AddReference(ulong bid) {
        ulong sequenceBid = bid & ~3UL;
        if (sequenceBid < 0x100 || (sequenceBid - 0x100) % 4 != 0) {
            throw new InvalidDataException("The PST block reference contains an invalid BID.");
        }
        long index = checked((long)((sequenceBid - 0x100) / 4));
        long offset = checked(index * RecordLength + 20);
        _writer.Flush();
        if (offset + sizeof(int) > _stream.Length) {
            throw new InvalidDataException("The PST block reference does not identify a written block.");
        }
        _stream.Position = offset;
        int read = _stream.Read(_referenceCountBuffer, 0, _referenceCountBuffer.Length);
        if (read != _referenceCountBuffer.Length) {
            throw new EndOfStreamException("The PST block journal reference count is truncated.");
        }
        int referenceCount = BitConverter.ToInt32(_referenceCountBuffer, 0);
        if (referenceCount < 1) {
            throw new InvalidDataException("The PST block journal contains an invalid reference count.");
        }
        _stream.Position = offset;
        _writer.Write(checked(referenceCount + 1));
    }

    internal PstWriterBlock Get(ulong bid) {
        ulong sequenceBid = bid & ~3UL;
        if (sequenceBid < 0x100 || (sequenceBid - 0x100) % 4 != 0) {
            throw new InvalidDataException("The PST block reference contains an invalid BID.");
        }
        long index = checked((long)((sequenceBid - 0x100) / 4));
        long offset = checked(index * RecordLength);
        _writer.Flush();
        if (offset + RecordLength > _stream.Length) {
            throw new InvalidDataException("The PST block reference does not identify a written block.");
        }
        _stream.Position = offset;
        using (var reader = new BinaryReader(_stream, Encoding.UTF8, leaveOpen: true)) {
            ulong storedBid = reader.ReadUInt64();
            long blockOffset = reader.ReadInt64();
            int length = reader.ReadInt32();
            int referenceCount = reader.ReadInt32();
            if ((storedBid & ~3UL) != sequenceBid) {
                throw new InvalidDataException("The PST block journal BID sequence is inconsistent.");
            }
            return new PstWriterBlock(storedBid, blockOffset, length, referenceCount);
        }
    }

    internal bool PrepareLegacyReferenceCountMigration() {
        _writer.Flush();
        bool hasZero = false;
        using (var input = new FileStream(_path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite, 64 * 1024, FileOptions.SequentialScan))
        using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
            while (input.Position < input.Length) {
                input.Position += 20;
                int value = reader.ReadInt32();
                if (value == 0) hasZero = true;
                else if (value < 0) throw new InvalidDataException(
                    "The PST block journal contains a negative reference count.");
            }
        }
        bool migrationPending = File.Exists(_referenceCountMigrationPath);
        if (!migrationPending && !hasZero) return false;
        if (!migrationPending) {
            using (var marker = new FileStream(_referenceCountMigrationPath, FileMode.CreateNew,
                FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough)) {
                marker.WriteByte(1);
                marker.Flush(flushToDisk: true);
            }
        }
        for (long index = 0; index < Count; index++) {
            _stream.Position = checked(index * RecordLength + 20);
            _writer.Write(1);
        }
        Flush(durable: true);
        return true;
    }

    internal void CompleteLegacyReferenceCountMigration() {
        Flush(durable: true);
        TryDelete(_referenceCountMigrationPath);
    }

    internal IEnumerable<PstWriterBlock> ReadAll() {
        _writer.Flush();
        using (var input = new FileStream(_path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite, 64 * 1024, FileOptions.SequentialScan))
        using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
            while (input.Position < input.Length) {
                ulong bid = reader.ReadUInt64();
                long offset = reader.ReadInt64();
                int length = reader.ReadInt32();
                int referenceCount = reader.ReadInt32();
                yield return new PstWriterBlock(bid, offset, length, referenceCount);
            }
        }
    }

    internal void Flush(bool durable) {
        _writer.Flush();
        _stream.Flush(durable);
    }

    internal void PreserveOnDispose() => _deleteOnDispose = false;

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _writer.Dispose();
        _stream.Dispose();
        if (_deleteOnDispose) {
            TryDelete(_path);
            TryDelete(_referenceCountMigrationPath);
        }
    }

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
