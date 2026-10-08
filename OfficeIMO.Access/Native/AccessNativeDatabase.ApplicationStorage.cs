using OfficeIMO.Core.Internal;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access;

internal sealed partial class AccessNativeDatabase {
    private Dictionary<string, AccessStorageStream> ReadApplicationStorage(CancellationToken cancellation) {
        var streams = new Dictionary<string, AccessStorageStream>(StringComparer.OrdinalIgnoreCase);
        if (_tables.TryGetValue("MSysAccessStorage", out var storage)) ReadAceStorage(storage, streams, cancellation);
        else if (_tables.TryGetValue("MSysAccessObjects", out var objects)) ReadJetStorage(objects, streams, cancellation);
        _document.ApplicationStreams = Array.AsReadOnly(streams.Values.ToArray());
        return streams;
    }

    private void ReadAceStorage(AccessNativeTable table, Dictionary<string, AccessStorageStream> streams, CancellationToken cancellation) {
        RequireFields(table, "Id", "Name", "ParentId", "Type", "Lv");
        var records = new Dictionary<int, StorageRecord>();
        using (var rows = new AccessNativeRowCursor(table, cancellation, rowLimit: MaxCatalogObjects)) while (rows.Read(cancellation)) {
            int id = Convert.ToInt32(RequiredField(table, rows, "Id", cancellation));
            var record = new StorageRecord(id, RequiredName(table, rows, "Name", cancellation),
                Convert.ToInt32(RequiredField(table, rows, "ParentId", cancellation)),
                Convert.ToInt32(RequiredField(table, rows, "Type", cancellation)), Field(table, rows, "Lv", cancellation) as byte[]);
            if (records.ContainsKey(id)) throw new InvalidDataException("Native application storage repeats a row identity.");
            records.Add(id, record);
        }
        foreach (var record in records.Values.Where(x => x.Type == 2)) {
            cancellation.ThrowIfCancellationRequested();
            var parts = new List<string> { record.Name }; var visited = new HashSet<int> { record.Id }; var parentId = record.ParentId;
            while (true) {
                if (visited.Count > MaxChainLength || !records.TryGetValue(parentId, out var parent) || parent.Type != 1)
                    throw new InvalidDataException("Native application storage has a missing or invalid parent.");
                if (parent.ParentId == parent.Id) {
                    if (parent.Name != "MSysAccessStorage_ROOT") throw new InvalidDataException("Native application stream does not reach its root.");
                    break;
                }
                if (!visited.Add(parent.Id)) throw new InvalidDataException("Native application storage has a parent cycle.");
                parts.Insert(0, parent.Name); parentId = parent.ParentId;
            }
            AddStorage(streams, string.Join("/", parts), record.Bytes ?? Array.Empty<byte>(), record.Id);
        }
    }

    private void ReadJetStorage(AccessNativeTable table, Dictionary<string, AccessStorageStream> streams, CancellationToken cancellation) {
        RequireFields(table, "ID", "Data"); var chunks = new SortedDictionary<int, byte[]>();
        using (var rows = new AccessNativeRowCursor(table, cancellation, rowLimit: MaxCatalogObjects)) while (rows.Read(cancellation)) {
            int id = Convert.ToInt32(RequiredField(table, rows, "ID", cancellation));
            byte[] chunk = (byte[])RequiredField(table, rows, "Data", cancellation);
            if (id > 0) {
                if (chunks.ContainsKey(id)) throw new InvalidDataException("Native Jet application carrier repeats a chunk identity.");
                chunks.Add(id, chunk);
            }
        }
        if (chunks.Count == 0) return;
        int size = checked(chunks.Values.Sum(x => x.Length));
        if (size > MaxMetadataBytes) throw new InvalidDataException("Native Jet application storage exceeds MaxMetadataBytes.");
        var bytes = new byte[size]; int offset = 0;
        foreach (var chunk in chunks.Values) { chunk.CopyTo(bytes, offset); offset += chunk.Length; }
        var limits = new OfficeCompoundReadOptions(MaxCatalogObjects, MaxCatalogObjects, MaxValueBytes, MaxMetadataBytes);
        if (!OfficeCompoundFileReader.TryRead(bytes, limits, cancellation, out OfficeCompoundFile? compound, out _) || compound == null) {
            AddApplicationDiagnostic("access.application.carrier-opaque", "The Jet application carrier is outside the qualified compound layout; its whole-file snapshot remains preserve-only.");
            return;
        }
        foreach (var pair in compound.Streams) AddStorage(streams, pair.Key, pair.Value, null);
    }

    private static void AddStorage(Dictionary<string, AccessStorageStream> streams, string path, byte[] bytes, int? nativeId) {
        if (path.Split('/').Any(x => x.Length == 0 || x == "." || x == ".." || x.IndexOf('\0') >= 0)
            || streams.ContainsKey(path)) throw new InvalidDataException("Native application storage contains an ambiguous path.");
        streams.Add(path, new AccessStorageStream(path, bytes, nativeId));
    }
    private void AddApplicationDiagnostic(string code, string message) => _document.Diagnostics = Array.AsReadOnly(
        _document.Diagnostics.Concat(new[] { new AccessDiagnostic(code, message) }).ToArray());
    private sealed class StorageRecord {
        internal StorageRecord(int id, string name, int parentId, int type, byte[]? bytes) { Id = id; Name = name; ParentId = parentId; Type = type; Bytes = bytes; }
        internal int Id, ParentId, Type; internal string Name; internal byte[]? Bytes;
    }
}
