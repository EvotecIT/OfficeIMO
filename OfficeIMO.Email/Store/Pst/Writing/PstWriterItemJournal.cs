using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Disk-backed folder-table rows with bounded external sorting.</summary>
internal sealed class PstWriterItemJournal : IDisposable {
    private const int RecordLength = 32;
    private const int MergeFanIn = 32;
    private readonly string _indexPath;
    private readonly string _payloadPath;
    private readonly FileStream _index;
    private readonly FileStream _payload;
    private readonly BinaryWriter _indexWriter;
    private bool _deleteOnDispose = true;
    private bool _disposed;

    internal PstWriterItemJournal(string pathPrefix, bool resume = false,
        long recordCount = 0, long payloadLength = 0) {
        _indexPath = string.Concat(pathPrefix, ".items");
        _payloadPath = string.Concat(pathPrefix, ".item-data");
        _index = new FileStream(_indexPath, resume ? FileMode.Open : FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.Read, 64 * 1024, FileOptions.SequentialScan);
        _payload = new FileStream(_payloadPath, resume ? FileMode.Open : FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.Read, 64 * 1024, FileOptions.SequentialScan);
        if (resume) {
            long indexLength = checked(recordCount * RecordLength);
            if (_index.Length < indexLength || _payload.Length < payloadLength) {
                throw new InvalidDataException("The PST item spool is truncated.");
            }
            _index.SetLength(indexLength);
            _payload.SetLength(payloadLength);
        }
        _indexWriter = new BinaryWriter(_index, Encoding.UTF8, leaveOpen: true);
    }

    internal int Count => checked((int)(_index.Length / RecordLength));
    internal long PayloadLength => _payload.Length;

    internal void Add(uint folderNid, uint nid, bool associated,
        IReadOnlyList<MapiProperty> properties) {
        long payloadOffset = _payload.Length;
        _payload.Position = payloadOffset;
        using (var buffer = new MemoryStream())
        using (var writer = new BinaryWriter(buffer, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(properties.Count);
            foreach (MapiProperty property in properties) WriteProperty(writer, property);
            writer.Flush();
            if (buffer.Length > int.MaxValue) throw new NotSupportedException("One PST table row payload is too large.");
            buffer.Position = 0;
            buffer.CopyTo(_payload);
        }
        int payloadLength = checked((int)(_payload.Length - payloadOffset));
        uint flags = associated ? 1U : 0U;
        if (IsUnread(properties)) flags |= 2U;
        _index.Position = _index.Length;
        WriteRecord(_indexWriter, new ItemRecord(folderNid, nid, flags,
            payloadOffset, payloadLength));
    }

    internal PstWriterItemSortedReader OpenSorted(int maximumRecordsInMemory) {
        if (maximumRecordsInMemory <= 0) throw new ArgumentOutOfRangeException(nameof(maximumRecordsInMemory));
        Flush(durable: false);
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
            string? run = runs.Count == 0 ? null : runs[0];
            var reader = new PstWriterItemSortedReader(run, _payloadPath);
            if (run != null) ownedRuns.Remove(run);
            foreach (string path in ownedRuns) TryDelete(path);
            return reader;
        } catch {
            foreach (string path in ownedRuns) TryDelete(path);
            throw;
        }
    }

    internal void Flush(bool durable) {
        _indexWriter.Flush();
        _payload.Flush(durable);
        _index.Flush(durable);
    }

    internal void PreserveOnDispose() => _deleteOnDispose = false;

    private List<string> CreateRuns(int maximumRecordsInMemory) {
        var runs = new List<string>();
        try {
            using (var input = new FileStream(_indexPath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite, 64 * 1024, FileOptions.SequentialScan))
            using (var reader = new BinaryReader(input, Encoding.UTF8, leaveOpen: false)) {
                while (input.Position < input.Length) {
                    var values = new List<ItemRecord>(maximumRecordsInMemory);
                    while (values.Count < maximumRecordsInMemory && input.Position < input.Length) {
                        values.Add(ReadRecord(reader));
                    }
                    values.Sort(Compare);
                    string run = NewRunPath();
                    try {
                        using (var output = new FileStream(run, FileMode.CreateNew, FileAccess.Write,
                            FileShare.Read, 64 * 1024, FileOptions.SequentialScan))
                        using (var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: false)) {
                            foreach (ItemRecord value in values) WriteRecord(writer, value);
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
                        if (selected == null || Compare(reader.Current, selected.Current) < 0) selected = reader;
                    }
                    if (selected == null) break;
                    WriteRecord(writer, selected.Current);
                    selected.MoveNext();
                }
            }
        } finally {
            foreach (RunReader reader in readers) reader.Dispose();
        }
    }

    private string NewRunPath() => string.Concat(_indexPath, ".sort.", Guid.NewGuid().ToString("N"));

    private static int Compare(ItemRecord left, ItemRecord right) {
        int result = left.FolderNid.CompareTo(right.FolderNid);
        if (result != 0) return result;
        result = left.Associated.CompareTo(right.Associated);
        return result != 0 ? result : left.Nid.CompareTo(right.Nid);
    }

    private static void WriteRecord(BinaryWriter writer, ItemRecord record) {
        writer.Write(record.FolderNid);
        writer.Write(record.Nid);
        writer.Write(record.Flags);
        writer.Write(0U);
        writer.Write(record.PayloadOffset);
        writer.Write(record.PayloadLength);
        writer.Write(0);
    }

    private static ItemRecord ReadRecord(BinaryReader reader) {
        uint folderNid = reader.ReadUInt32();
        uint nid = reader.ReadUInt32();
        uint flags = reader.ReadUInt32();
        reader.ReadUInt32();
        long payloadOffset = reader.ReadInt64();
        int payloadLength = reader.ReadInt32();
        reader.ReadInt32();
        return new ItemRecord(folderNid, nid, flags, payloadOffset, payloadLength);
    }

    private static void WriteProperty(BinaryWriter writer, MapiProperty property) {
        writer.Write(property.PropertyId);
        writer.Write((ushort)property.PropertyType);
        writer.Write(property.Flags);
        byte[] value;
        if (PstPropertyValueWriter.IsInline(property.PropertyType)) {
            value = BitConverter.GetBytes(PstPropertyValueWriter.EncodeInline(property));
        } else {
            value = PstPropertyValueWriter.EncodeVariable(property, 65001);
        }
        writer.Write(value.Length);
        writer.Write(value);
    }

    private static MapiProperty ReadProperty(BinaryReader reader) {
        ushort id = reader.ReadUInt16();
        var type = (MapiPropertyType)reader.ReadUInt16();
        uint flags = reader.ReadUInt32();
        int length = reader.ReadInt32();
        if (length < 0 || length > 64 * 1024 * 1024) {
            throw new InvalidDataException("A spooled PST table value has an invalid length.");
        }
        byte[] value = reader.ReadBytes(length);
        if (value.Length != length) throw new EndOfStreamException("A spooled PST table value is truncated.");
        if (!PstPropertyValueWriter.IsInline(type)) {
            return new MapiProperty(id, type, null, flags) { RawData = value };
        }
        if (value.Length != 4) throw new InvalidDataException("A spooled inline MAPI value is invalid.");
        uint inline = BitConverter.ToUInt32(value, 0);
        object? decoded;
        switch (type) {
            case MapiPropertyType.Null: decoded = null; break;
            case MapiPropertyType.Integer16: decoded = unchecked((short)inline); break;
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode: decoded = unchecked((int)inline); break;
            case MapiPropertyType.Floating32: decoded = BitConverter.ToSingle(value, 0); break;
            case MapiPropertyType.Boolean: decoded = inline != 0; break;
            default: decoded = unchecked((int)inline); break;
        }
        return new MapiProperty(id, type, decoded, flags);
    }

    private static bool IsUnread(IEnumerable<MapiProperty> properties) {
        int? flags = properties.GetNullableMapiValue(MapiKnownProperties.PidTag.MessageFlags);
        return !flags.HasValue || (flags.Value & 1) == 0;
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _indexWriter.Dispose();
        _index.Dispose();
        _payload.Dispose();
        if (_deleteOnDispose) {
            TryDelete(_indexPath);
            TryDelete(_payloadPath);
        }
    }

    private static void TryDelete(string path) {
        try { if (path != null && File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    internal sealed class PstWriterItemSortedReader : IDisposable {
        private readonly string? _runPath;
        private readonly FileStream? _run;
        private readonly BinaryReader? _reader;
        private readonly FileStream _payload;
        private ItemRecord? _current;
        private uint _lastFolder;
        private bool _lastAssociated;
        private bool _hasRequest;

        internal PstWriterItemSortedReader(string? runPath, string payloadPath) {
            _runPath = runPath;
            if (runPath != null) {
                _run = new FileStream(runPath, FileMode.Open, FileAccess.Read,
                    FileShare.Read, 64 * 1024, FileOptions.SequentialScan);
                _reader = new BinaryReader(_run, Encoding.UTF8, leaveOpen: true);
                MoveNext();
            }
            _payload = new FileStream(payloadPath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite, 64 * 1024, FileOptions.RandomAccess);
        }

        internal IEnumerable<PstWriterTableRow> ReadRows(uint folderNid, bool associated) {
            if (_hasRequest && (folderNid < _lastFolder ||
                (folderNid == _lastFolder && associated.CompareTo(_lastAssociated) <= 0))) {
                throw new InvalidOperationException("PST item spool groups must be requested in ascending order.");
            }
            _hasRequest = true;
            _lastFolder = folderNid;
            _lastAssociated = associated;
            while (_current.HasValue && _current.Value.FolderNid == folderNid &&
                _current.Value.Associated == associated) {
                ItemRecord record = _current.Value;
                yield return ReadRow(record);
                MoveNext();
            }
        }

        internal bool IsExhausted => !_current.HasValue;

        private PstWriterTableRow ReadRow(ItemRecord record) {
            _payload.Position = record.PayloadOffset;
            var bytes = new byte[record.PayloadLength];
            int total = 0;
            while (total < bytes.Length) {
                int read = _payload.Read(bytes, total, bytes.Length - total);
                if (read == 0) throw new EndOfStreamException("A spooled PST table row is truncated.");
                total += read;
            }
            using (var buffer = new MemoryStream(bytes, writable: false))
            using (var reader = new BinaryReader(buffer, Encoding.UTF8, leaveOpen: false)) {
                int count = reader.ReadInt32();
                if (count < 0 || count > 512) throw new InvalidDataException("A spooled PST table row has an invalid property count.");
                var properties = new MapiProperty[count];
                for (int index = 0; index < count; index++) properties[index] = ReadProperty(reader);
                return new PstWriterTableRow(record.Nid, properties);
            }
        }

        private void MoveNext() {
            if (_run == null || _reader == null || _run.Position >= _run.Length) {
                _current = null;
                return;
            }
            _current = ReadRecord(_reader);
        }

        public void Dispose() {
            _reader?.Dispose();
            _run?.Dispose();
            _payload.Dispose();
            if (_runPath != null) TryDelete(_runPath);
        }
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
        internal ItemRecord Current { get; private set; }
        internal void MoveNext() {
            HasValue = _stream.Position < _stream.Length;
            if (HasValue) Current = ReadRecord(_reader);
        }
        public void Dispose() { _reader.Dispose(); _stream.Dispose(); }
    }

    private readonly struct ItemRecord {
        internal ItemRecord(uint folderNid, uint nid, uint flags,
            long payloadOffset, int payloadLength) {
            FolderNid = folderNid;
            Nid = nid;
            Flags = flags;
            PayloadOffset = payloadOffset;
            PayloadLength = payloadLength;
        }
        internal uint FolderNid { get; }
        internal uint Nid { get; }
        internal uint Flags { get; }
        internal bool Associated => (Flags & 1) != 0;
        internal long PayloadOffset { get; }
        internal int PayloadLength { get; }
    }
}
