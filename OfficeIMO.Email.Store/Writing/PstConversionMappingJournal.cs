namespace OfficeIMO.Email.Store;

/// <summary>Sequential, delete-on-close source-to-destination verification mappings.</summary>
internal sealed class PstConversionMappingJournal : IDisposable {
    private const int MaximumStringBytes = 16 * 1024 * 1024;
    private readonly string _path;
    private readonly FileStream _stream;
    private readonly BinaryWriter _writer;
    private bool _reading;
    private bool _disposed;

    internal PstConversionMappingJournal(string destinationPath) {
        _path = string.Concat(destinationPath, ".", Guid.NewGuid().ToString("N"), ".verify-map.tmp");
        _stream = new FileStream(_path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read,
            64 * 1024, FileOptions.SequentialScan | FileOptions.DeleteOnClose);
        _writer = new BinaryWriter(_stream, Encoding.UTF8, leaveOpen: true);
    }

    internal int Count { get; private set; }

    internal long Length => _stream.Length;

    internal void Add(int ordinal, EmailStoreItemReference source, string destinationFolderId,
        string destinationItemId) {
        if (_reading) throw new InvalidOperationException("The conversion mapping journal is already being read.");
        _writer.Write(ordinal);
        byte flags = 0;
        if (source.IsAssociated) flags |= 1;
        if (source.IsOrphaned) flags |= 2;
        _writer.Write(flags);
        _writer.Write(source.Id);
        _writer.Write(source.FolderId);
        _writer.Write(destinationFolderId);
        _writer.Write(destinationItemId);
        Count = checked(Count + 1);
    }

    internal IEnumerable<PstConversionItemMap> ReadAll() {
        if (_reading) throw new InvalidOperationException("The conversion mapping journal can be enumerated once.");
        _reading = true;
        _writer.Flush();
        _stream.Flush(flushToDisk: false);
        _stream.Position = 0;
        using (var reader = new BinaryReader(_stream, Encoding.UTF8, leaveOpen: true)) {
            for (int index = 0; index < Count; index++) {
                int ordinal = reader.ReadInt32();
                byte flags = reader.ReadByte();
                string sourceId = ReadBoundedString(reader);
                string sourceFolderId = ReadBoundedString(reader);
                string destinationFolderId = ReadBoundedString(reader);
                string destinationItemId = ReadBoundedString(reader);
                var source = new EmailStoreItemReference(sourceId, sourceFolderId,
                    (flags & 1) != 0, (flags & 2) != 0);
                yield return new PstConversionItemMap(ordinal, source,
                    destinationFolderId, destinationItemId);
            }
            if (_stream.Position != _stream.Length) {
                throw new InvalidDataException("The conversion mapping journal has trailing data.");
            }
        }
    }

    private static string ReadBoundedString(BinaryReader reader) {
        long start = reader.BaseStream.Position;
        string value = reader.ReadString();
        if (reader.BaseStream.Position - start > MaximumStringBytes) {
            throw new InvalidDataException("A conversion mapping identifier exceeds the configured journal bound.");
        }
        return value;
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _writer.Dispose();
        _stream.Dispose();
        try { if (File.Exists(_path)) File.Delete(_path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}

internal sealed class PstConversionItemMap {
    internal PstConversionItemMap(int ordinal, EmailStoreItemReference source,
        string destinationFolderId, string destinationItemId) {
        Ordinal = ordinal;
        Source = source;
        DestinationFolderId = destinationFolderId;
        DestinationItemId = destinationItemId;
    }
    internal int Ordinal { get; }
    internal EmailStoreItemReference Source { get; }
    internal string DestinationFolderId { get; }
    internal string DestinationItemId { get; }
}
