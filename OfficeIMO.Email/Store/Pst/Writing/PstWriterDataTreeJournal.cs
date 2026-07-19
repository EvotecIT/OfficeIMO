namespace OfficeIMO.Email.Store;

/// <summary>Temporary fixed-record data-tree level used to bound large attachment indexing.</summary>
internal sealed class PstWriterDataTreeJournal : IDisposable {
    private const int RecordLength = 16;
    private readonly string _path;
    private readonly FileStream _stream;
    private readonly BinaryWriter _writer;
    private bool _disposed;

    internal PstWriterDataTreeJournal(string path) {
        _path = path;
        _stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.Read, 32 * 1024, FileOptions.SequentialScan | FileOptions.DeleteOnClose);
        _writer = new BinaryWriter(_stream, Encoding.UTF8, leaveOpen: true);
    }

    internal int Count => checked((int)(_stream.Length / RecordLength));

    internal void Add(ulong bid, uint length) {
        _stream.Position = _stream.Length;
        _writer.Write(bid);
        _writer.Write(length);
        _writer.Write(0U);
    }

    internal IEnumerable<PstWriterDataTreeReference> ReadAll() {
        _writer.Flush();
        using (var input = new FileStream(_path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete, 32 * 1024, FileOptions.SequentialScan))
        using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
            while (input.Position < input.Length) {
                ulong bid = reader.ReadUInt64();
                uint length = reader.ReadUInt32();
                reader.ReadUInt32();
                yield return new PstWriterDataTreeReference(bid, length);
            }
        }
    }

    internal PstWriterDataTreeReference ReadSingle() => ReadAll().Single();

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _writer.Dispose();
        _stream.Dispose();
    }
}

internal readonly struct PstWriterDataTreeReference {
    internal PstWriterDataTreeReference(ulong bid, uint length) {
        Bid = bid;
        Length = length;
    }
    internal ulong Bid { get; }
    internal uint Length { get; }
}
