namespace OfficeIMO.Email.Store;

/// <summary>Temporary fixed-record B-tree level used while constructing large indexes.</summary>
internal sealed class PstWriterPageReferenceJournal : IDisposable {
    private const int RecordLength = 24;
    private readonly string _path;
    private readonly FileStream _stream;
    private readonly BinaryWriter _writer;
    private bool _disposed;

    internal PstWriterPageReferenceJournal(string path) {
        _path = path;
        _stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.Read, 16 * 1024, FileOptions.SequentialScan | FileOptions.DeleteOnClose);
        _writer = new BinaryWriter(_stream, Encoding.UTF8, leaveOpen: true);
    }

    internal int Count => checked((int)(_stream.Length / RecordLength));

    internal void Add(PstWriterPageReference value) {
        _stream.Position = _stream.Length;
        _writer.Write(value.Key);
        _writer.Write(value.Bid);
        _writer.Write(value.Offset);
    }

    internal IEnumerable<PstWriterPageReference> ReadAll() {
        _writer.Flush();
        using (var input = new FileStream(_path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete, 16 * 1024, FileOptions.SequentialScan))
        using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
            while (input.Position < input.Length) {
                yield return new PstWriterPageReference(
                    reader.ReadUInt64(), reader.ReadUInt64(), reader.ReadInt64());
            }
        }
    }

    internal PstWriterPageReference ReadSingle() => ReadAll().Single();

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _writer.Dispose();
        _stream.Dispose();
    }
}
