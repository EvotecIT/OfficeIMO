namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        private AccessNativeWriter BuildAceApplicationMutation(AccessNativeTable table, IReadOnlyDictionary<string, byte[]> replacements,
            ISet<string> removals, long maximumBytes, CancellationToken cancellation) {
            AccessNativeColumn payload = table.Columns.Single(x => x.Name.Equals("Lv", StringComparison.OrdinalIgnoreCase));
            var records = new Dictionary<int, (StorageRecord Record, uint Pointer)>();
            using (var cursor = new AccessNativeRowCursor(table, cancellation, rowLimit: MaxCatalogObjects)) while (cursor.Read(cancellation)) {
                int id = Convert.ToInt32(RequiredField(table, cursor, "Id", cancellation));
                records.Add(id, (new StorageRecord(id, RequiredName(table, cursor, "Name", cancellation),
                    Convert.ToInt32(RequiredField(table, cursor, "ParentId", cancellation)),
                    Convert.ToInt32(RequiredField(table, cursor, "Type", cancellation)), null), cursor.CurrentPointer));
            }
            var paths = ReadAceStoragePaths(records.ToDictionary(x => x.Key, x => x.Value.Record), cancellation);
            var kinds = records.ToDictionary(x => x.Key, x => x.Value.Record.Type);
            var values = new Dictionary<uint, byte[]>(); var deleted = new HashSet<uint>(); var additions = new List<object?[]>();
            int nextId = records.Keys.Max();
            foreach (string path in removals) {
                if (!paths.TryGetValue(path, out int id)) continue;
                if (records[id].Record.Type != 2) throw new NotSupportedException("Application mutation removes only explicit stream identities.");
                deleted.Add(records[id].Pointer);
            }
            foreach (var pair in replacements) {
                cancellation.ThrowIfCancellationRequested();
                if (paths.TryGetValue(pair.Key, out int id)) {
                    if (records[id].Record.Type != 2) throw new InvalidDataException("A new application stream collides with an existing storage.");
                    values.Add(records[id].Pointer, pair.Value); continue;
                }
                string[] parts = pair.Key.Split('/'); string parent = string.Empty;
                for (int index = 0; index < parts.Length; index++) {
                    string path = parent.Length == 0 ? parts[index] : parent + "/" + parts[index];
                    if (!paths.TryGetValue(path, out int child)) {
                        child = checked(++nextId); object?[] row = new object?[table.Columns.Count];
                        void Set(string field, object? value) => row[table.Columns.FindIndex(x => x.Name.Equals(field, StringComparison.OrdinalIgnoreCase))] = value;
                        Set("Id", child); Set("Name", parts[index]); Set("ParentId", paths[parent]); Set("Type", index == parts.Length - 1 ? 2 : 1);
                        Set("DateCreate", _document.CreatedAt.ToOADate()); Set("DateUpdate", _document.CreatedAt.ToOADate());
                        Set("Lv", index == parts.Length - 1 ? pair.Value : null);
                        additions.Add(row); paths.Add(path, child); kinds.Add(child, index == parts.Length - 1 ? 2 : 1);
                    } else if (index < parts.Length - 1 && kinds[child] != 1)
                        throw new InvalidDataException("A new application storage collides with a stream.");
                    parent = path;
                }
            }
            AccessNativeWriter writer = AccessNativeWriter.BuildBinaryReplacements(this, table, payload, values, maximumBytes, cancellation);
            writer.MutateApplicationRows(this, table, additions, deleted);
            return writer;
        }
    }
}
