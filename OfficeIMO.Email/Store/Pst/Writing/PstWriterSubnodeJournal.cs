namespace OfficeIMO.Email.Store;

/// <summary>Append-only sorted subnode source kept off the managed heap.</summary>
internal sealed class PstWriterSubnodeJournal : IDisposable {
    private const int RecordLength = 24;
    private readonly string _path;
    private readonly FileStream _stream;
    private readonly BinaryWriter _writer;
    private uint _lastNid;
    private bool _disposed;

    internal PstWriterSubnodeJournal(string path) {
        _path = path;
        _stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.Read, 32 * 1024, FileOptions.SequentialScan);
        _writer = new BinaryWriter(_stream, Encoding.UTF8, leaveOpen: true);
    }

    internal int Count => checked((int)(_stream.Length / RecordLength));

    internal void Add(PstWriterSubnode value) {
        if (Count > 0 && value.Nid <= _lastNid) {
            throw new InvalidDataException("PST subnode journal entries must have unique ascending NIDs.");
        }
        _stream.Position = _stream.Length;
        _writer.Write(value.Nid);
        _writer.Write(0U);
        _writer.Write(value.DataBid);
        _writer.Write(value.SubnodeBid);
        _lastNid = value.Nid;
    }

    internal IEnumerable<PstWriterSubnode> ReadAll() {
        _writer.Flush();
        using (var input = new FileStream(_path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite, 32 * 1024, FileOptions.SequentialScan))
        using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
            while (input.Position < input.Length) {
                uint nid = reader.ReadUInt32();
                reader.ReadUInt32();
                yield return new PstWriterSubnode(nid, reader.ReadUInt64(), reader.ReadUInt64());
            }
        }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _writer.Dispose();
        _stream.Dispose();
        TryDelete(_path);
    }

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
