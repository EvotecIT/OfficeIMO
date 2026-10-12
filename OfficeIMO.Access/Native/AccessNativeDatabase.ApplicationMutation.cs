using OfficeIMO.Core.Internal;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        /// <summary>Plans bounded application-stream changes while retaining all surviving native storage and row identities.</summary>
        internal AccessNativeWriter BuildApplicationStreamReplacement(IReadOnlyDictionary<string, byte[]> replacements,
            long maximumBytes, CancellationToken cancellation, ISet<string>? removals = null) {
            cancellation.ThrowIfCancellationRequested();
            var byPath = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            var deleted = new HashSet<string>(removals ?? new HashSet<string>(), StringComparer.OrdinalIgnoreCase);
            long total = 0;
            foreach (KeyValuePair<string, byte[]> replacement in replacements) {
                cancellation.ThrowIfCancellationRequested();
                _applicationStreams.TryGetValue(replacement.Key, out AccessStorageStream? original);
                byte[] value = replacement.Value ?? throw new ArgumentException("Application replacement payload cannot be null.", nameof(replacements));
                if (value.Length > MaxMetadataBytes - total) throw new InvalidDataException("Application replacements exceed MaxMetadataBytes.");
                total += value.Length;
                string path = original?.Path ?? replacement.Key;
                if (path.Split('/').Any(x => x.Length == 0 || x == "." || x == ".." || x.IndexOf('\0') >= 0) || deleted.Contains(path))
                    throw new ArgumentException("Application changes contain invalid or conflicting paths.", nameof(replacements));
                if (byPath.ContainsKey(path)) throw new ArgumentException("Application replacements contain ambiguous path aliases.", nameof(replacements));
                byPath.Add(path, value);
            }
            if (_tables.TryGetValue("MSysAccessStorage", out AccessNativeTable? ace)) {
                return BuildAceApplicationMutation(ace, byPath, deleted, maximumBytes, cancellation);
            }
            if (!_tables.TryGetValue("MSysAccessObjects", out AccessNativeTable? jet) || _jetApplicationCompound == null)
                throw new NotSupportedException("The native Jet application carrier is not decoded.");
            byte[] compound = OfficeCompoundFileWriter.Rewrite(_jetApplicationCompound, byPath, deleted,
                maxOutputBytes: MaxMetadataBytes, cancellationToken: cancellation);
            AccessNativeColumn data = jet.Columns.Single(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, "Data"));
            var chunks = new SortedDictionary<int, (uint Pointer, int Length)>();
            using (var cursor = new AccessNativeRowCursor(jet, cancellation, rowLimit: MaxCatalogObjects)) {
                while (cursor.Read(cancellation)) {
                    int id = Convert.ToInt32(RequiredField(jet, cursor, "ID", cancellation));
                    if (id <= 0) continue;
                    int ordinal = jet.Columns.IndexOf(data);
                    using Stream stream = cursor.OpenBinary(ordinal, cancellation);
                    chunks.Add(id, (cursor.CurrentPointer, checked((int)stream.Length)));
                }
            }
            int capacity = checked(chunks.Values.Sum(x => x.Length));
            var jetValues = new Dictionary<uint, byte[]>(); int offset = 0;
            foreach ((uint Pointer, int Length) chunk in chunks.Values) {
                cancellation.ThrowIfCancellationRequested(); byte[] value = new byte[chunk.Length];
                int count = Math.Min(chunk.Length, compound.Length - offset);
                if (count > 0) Buffer.BlockCopy(compound, offset, value, 0, count);
                offset += count; jetValues.Add(chunk.Pointer, value);
            }
            AccessNativeWriter writer = AccessNativeWriter.BuildBinaryReplacements(this, jet, data, jetValues, maximumBytes, cancellation);
            if (compound.Length > capacity) {
                if (chunks.Count == 0 || chunks.Keys.First() != 1 || chunks.Keys.Last() != chunks.Count || chunks.Values.Any(x => x.Length != 3992)
                    || data.Variable || data.Size != 3992)
                    throw new NotSupportedException("Jet application growth requires the qualified consecutive fixed-width carrier.");
                var additions = new List<object?[]>(); int id = chunks.Count;
                while (offset < compound.Length) {
                    cancellation.ThrowIfCancellationRequested(); byte[] chunk = new byte[3992]; int length = Math.Min(chunk.Length, compound.Length - offset);
                    Buffer.BlockCopy(compound, offset, chunk, 0, length); offset += length;
                    object?[] row = new object?[jet.Columns.Count]; row[jet.Columns.IndexOf(data)] = chunk;
                    row[jet.Columns.FindIndex(x => x.Name.Equals("ID", StringComparison.OrdinalIgnoreCase))] = checked(++id); additions.Add(row);
                }
                writer.MutateApplicationRows(this, jet, additions, new HashSet<uint>());
            }
            return writer;
        }
    }
}
