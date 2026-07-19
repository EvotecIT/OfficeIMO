namespace OfficeIMO.Email.Store;

/// <summary>Disk-backed NBT source with bounded external sorting.</summary>
internal sealed class PstWriterNodeJournal : IDisposable {
    private const int RecordLength = 24;
    private const int MergeFanIn = 32;
    private readonly string _path;
    private readonly FileStream _stream;
    private readonly BinaryWriter _writer;
    private readonly uint[] _maximumIndexes = Enumerable.Repeat(0x400U, 32).ToArray();
    private bool _deleteOnDispose = true;
    private bool _disposed;

    internal PstWriterNodeJournal(string path, bool resume = false, long recordCount = 0) {
        _path = path;
        _stream = new FileStream(path, resume ? FileMode.Open : FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.Read, 64 * 1024, FileOptions.SequentialScan);
        if (resume) {
            long length = checked(recordCount * RecordLength);
            if (_stream.Length < length) throw new InvalidDataException("The PST node journal is truncated.");
            _stream.SetLength(length);
        }
        _writer = new BinaryWriter(_stream, Encoding.UTF8, leaveOpen: true);
        _maximumIndexes[2] = 0x400;
        _maximumIndexes[3] = 0x4000;
        _maximumIndexes[4] = 0x10000;
        _maximumIndexes[8] = 0x8000;
        if (resume) RebuildMaximumIndexes();
    }

    internal int Count => checked((int)(_stream.Length / RecordLength));
    internal IReadOnlyList<uint> MaximumIndexes => _maximumIndexes;

    internal void Add(PstWriterNode node) {
        _stream.Position = _stream.Length;
        Write(_writer, node);
        int type = checked((int)(node.Nid & 0x1F));
        uint index = node.Nid >> 5;
        if (index > _maximumIndexes[type]) _maximumIndexes[type] = index;
    }

    internal IEnumerable<PstWriterNode> ReadSorted(int maximumRecordsInMemory) {
        if (maximumRecordsInMemory <= 0) throw new ArgumentOutOfRangeException(nameof(maximumRecordsInMemory));
        _writer.Flush();
        List<string> runs = CreateRuns(maximumRecordsInMemory);
        var ownedRuns = new HashSet<string>(runs, StringComparer.Ordinal);
        try {
            while (runs.Count > 1) {
                var merged = new List<string>((runs.Count + MergeFanIn - 1) / MergeFanIn);
                for (int offset = 0; offset < runs.Count; offset += MergeFanIn) {
                    string[] group = runs.Skip(offset).Take(MergeFanIn).ToArray();
                    string output = NewRunPath();
                    ownedRuns.Add(output);
                    try { MergeRuns(group, output); }
                    catch { TryDelete(output); throw; }
                    merged.Add(output);
                    foreach (string path in group) TryDelete(path);
                }
                runs = merged;
            }
            if (runs.Count == 0) yield break;
            foreach (PstWriterNode node in ReadRun(runs[0])) yield return node;
        } finally {
            foreach (string path in ownedRuns) TryDelete(path);
        }
    }

    internal void Flush(bool durable) {
        _writer.Flush();
        _stream.Flush(durable);
    }

    internal void PreserveOnDispose() => _deleteOnDispose = false;

    private void RebuildMaximumIndexes() {
        foreach (PstWriterNode node in ReadAllUnsorted()) {
            int type = checked((int)(node.Nid & 0x1F));
            _maximumIndexes[type] = Math.Max(_maximumIndexes[type], node.Nid >> 5);
        }
    }

    private IEnumerable<PstWriterNode> ReadAllUnsorted() {
        using (var input = new FileStream(_path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite, 64 * 1024, FileOptions.SequentialScan))
        using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
            while (input.Position < input.Length) yield return Read(reader);
        }
    }

    private List<string> CreateRuns(int maximumRecordsInMemory) {
        var runs = new List<string>();
        try {
            using (var input = new FileStream(_path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite, 64 * 1024, FileOptions.SequentialScan))
            using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
                while (input.Position < input.Length) {
                    var values = new List<PstWriterNode>(maximumRecordsInMemory);
                    while (values.Count < maximumRecordsInMemory && input.Position < input.Length) {
                        values.Add(Read(reader));
                    }
                    values.Sort((left, right) => left.Nid.CompareTo(right.Nid));
                    string run = NewRunPath();
                    try {
                        using (var output = new FileStream(run, FileMode.CreateNew, FileAccess.Write,
                            FileShare.Read, 64 * 1024, FileOptions.SequentialScan))
                        using (var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: false)) {
                            foreach (PstWriterNode value in values) Write(writer, value);
                        }
                        runs.Add(run);
                    } catch {
                        TryDelete(run);
                        throw;
                    }
                }
            }
            return runs;
        } catch {
            foreach (string run in runs) TryDelete(run);
            throw;
        }
    }

    private static void MergeRuns(IReadOnlyList<string> inputs, string outputPath) {
        var readers = new List<RunReader>(inputs.Count);
        try {
            foreach (string path in inputs) readers.Add(new RunReader(path));
            using (var output = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write,
                FileShare.Read, 64 * 1024, FileOptions.SequentialScan))
            using (var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: false)) {
                while (true) {
                    RunReader? selected = null;
                    foreach (RunReader reader in readers) {
                        if (!reader.HasValue) continue;
                        if (selected == null || reader.Current!.Nid < selected.Current!.Nid) selected = reader;
                    }
                    if (selected == null) break;
                    Write(writer, selected.Current!);
                    selected.MoveNext();
                }
            }
        } finally {
            foreach (RunReader reader in readers) reader.Dispose();
        }
    }

    private static IEnumerable<PstWriterNode> ReadRun(string path) {
        using (var input = new FileStream(path, FileMode.Open, FileAccess.Read,
            FileShare.Read, 64 * 1024, FileOptions.SequentialScan))
        using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
            while (input.Position < input.Length) yield return Read(reader);
        }
    }

    private string NewRunPath() => string.Concat(_path, ".sort.", Guid.NewGuid().ToString("N"));

    private static void Write(BinaryWriter writer, PstWriterNode node) {
        writer.Write(node.Nid);
        writer.Write(node.ParentNid);
        writer.Write(node.DataBid);
        writer.Write(node.SubnodeBid);
    }

    private static PstWriterNode Read(BinaryReader reader) => new PstWriterNode(
        reader.ReadUInt32(), reader.ReadUInt32(), reader.ReadUInt64(), reader.ReadUInt64());

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

    private sealed class RunReader : IDisposable {
        private readonly FileStream _stream;
        private readonly BinaryReader _reader;
        internal RunReader(string path) {
            _stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.Read, 16 * 1024, FileOptions.SequentialScan);
            _reader = new BinaryReader(_stream, Encoding.UTF8, leaveOpen: true);
            MoveNext();
        }
        internal bool HasValue { get; private set; }
        internal PstWriterNode? Current { get; private set; }
        internal void MoveNext() {
            HasValue = _stream.Position < _stream.Length;
            Current = HasValue ? Read(_reader) : null;
        }
        public void Dispose() { _reader.Dispose(); _stream.Dispose(); }
    }
}
