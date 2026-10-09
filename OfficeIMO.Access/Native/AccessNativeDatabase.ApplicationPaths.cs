using OfficeIMO.Core.Internal;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        /// <summary>Validates the complete native ACE storage hierarchy, including empty storages.</summary>
        private Dictionary<string, int> ReadAceStoragePaths(IReadOnlyDictionary<int, StorageRecord> records, CancellationToken cancellation) {
            var paths = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (StorageRecord record in records.Values) {
                cancellation.ThrowIfCancellationRequested();
                if (record.Type != 1 && record.Type != 2 || record.Name.Length == 0 || record.Name.IndexOfAny(new[] { '/', '\\', '\0' }) >= 0 || record.Name == "." || record.Name == "..")
                    throw new InvalidDataException("An ACE application storage has an invalid type or path component.");
                StorageRecord current = record;
                if (current.ParentId == current.Id) {
                    if (current.Type != 1 || current.Name != "MSysAccessStorage_ROOT" && current.Name != "MSysAccessStorage_SCRATCH")
                        throw new InvalidDataException("An ACE application root is invalid.");
                    Add(current.Name == "MSysAccessStorage_ROOT" ? string.Empty : current.Name, current.Id); continue;
                }
                var parts = new List<string> { current.Name }; var seen = new HashSet<int> { current.Id };
                while (true) {
                    cancellation.ThrowIfCancellationRequested();
                    if (seen.Count >= MaxChainLength || !records.TryGetValue(current.ParentId, out StorageRecord? parent) || parent.Type != 1 || !seen.Add(parent.Id))
                        throw new InvalidDataException("An ACE application storage has an invalid parent graph.");
                    current = parent;
                    if (current.ParentId == current.Id) {
                        if (current.Name != "MSysAccessStorage_ROOT") throw new InvalidDataException("An ACE application storage does not reach its root.");
                        break;
                    }
                    parts.Insert(0, current.Name);
                }
                Add(string.Join("/", parts), record.Id);
            }
            return paths;
            void Add(string path, int id) {
                if (paths.ContainsKey(path)) throw new InvalidDataException("ACE application storage contains ambiguous identities.");
                paths.Add(path, id);
            }
        }

        /// <summary>Detaches the stored VBA subtree without losing empty storage identities or compound metadata.</summary>
        internal byte[] DetachVbaProject(long maximumBytes, CancellationToken cancellation) {
            const string prefix = "VBA/VBAProject/";
            var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            long total = 0;
            foreach (AccessStorageStream stream in _applicationStreams.Values) {
                cancellation.ThrowIfCancellationRequested();
                if (!stream.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                if (stream.Payload.Length > maximumBytes - total) throw new InvalidDataException("The Access VBA streams exceed the configured project byte limit.");
                total += stream.Payload.Length;
                streams.Add(stream.Path.Substring(prefix.Length), stream.Payload.GetBytes());
            }
            if (streams.Count == 0) throw new InvalidOperationException("The Access database has no stored VBA project.");
            OfficeCompoundFileEntry? originalRoot = _applicationEntries.SingleOrDefault(x => x.Path.Equals(prefix.TrimEnd('/'), StringComparison.OrdinalIgnoreCase));
            if (originalRoot == null || !originalRoot.IsStorage || originalRoot.IsFallback) throw new InvalidDataException("The Access VBA project has no qualified storage root.");
            var root = new OfficeCompoundFileEntry("Root Entry", string.Empty, 5, 0, classId: originalRoot.ClassId, stateBits: originalRoot.StateBits,
                creationTime: originalRoot.CreationTime, modifiedTime: originalRoot.ModifiedTime, directoryOrder: originalRoot.DirectoryOrder);
            var entries = new List<OfficeCompoundFileEntry>();
            foreach (OfficeCompoundFileEntry entry in _applicationEntries.Where(x => x.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))) {
                cancellation.ThrowIfCancellationRequested();
                if (entry.IsFallback) throw new InvalidDataException("The Access VBA storage has an unreachable compound identity.");
                entries.Add(new OfficeCompoundFileEntry(entry.Name, entry.Path.Substring(prefix.Length), entry.ObjectType, entry.Size,
                    classId: entry.ClassId, stateBits: entry.StateBits, creationTime: entry.CreationTime, modifiedTime: entry.ModifiedTime, directoryOrder: entry.DirectoryOrder));
            }
            return OfficeCompoundFileWriter.Rewrite(new OfficeCompoundFile(streams, entries, root), new Dictionary<string, byte[]>(),
                maxOutputBytes: maximumBytes, cancellationToken: cancellation);
        }
    }
}
