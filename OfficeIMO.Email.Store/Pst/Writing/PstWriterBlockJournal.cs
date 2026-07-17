namespace OfficeIMO.Email.Store;

/// <summary>Append-only fixed-record block index kept off the managed heap.</summary>
internal sealed class PstWriterBlockJournal : IDisposable {
    private const int RecordLength = 24;
    private readonly string _path;
    private readonly FileStream _stream;
    private readonly BinaryWriter _writer;
    private bool _deleteOnDispose = true;
    private bool _disposed;

    internal PstWriterBlockJournal(string path, bool resume = false, long recordCount = 0) {
        _path = path;
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
        _writer.Write(0);
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
                reader.ReadInt32();
                yield return new PstWriterBlock(bid, offset, length);
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
        if (_deleteOnDispose) TryDelete(_path);
    }

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
