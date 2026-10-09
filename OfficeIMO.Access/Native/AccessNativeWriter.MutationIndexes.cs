using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        /// <summary>Rebuilds only the affected index trees, retaining encoded keys for all original rows.</summary>
        private void UpdateMutationIndexes(AccessNativeDatabase database, AccessNativeTable table,
            IReadOnlyList<object?[]> additions, IReadOnlyList<uint> rowIds, ISet<uint> removals, Column[] columns,
            IReadOnlyDictionary<uint, object?[]>? updates = null, long? rowLimit = null) {
            Table shape = new Table(table.DefinitionPage, table.Name, columns, additions.ToArray());
            var live = new HashSet<uint>(); var physical = new HashSet<uint>();
            using (var rows = new AccessNativeRowCursor(table, _cancellation, rowLimit: rowLimit ?? database.MaxCatalogObjects)) while (rows.Read(_cancellation)) {
                live.Add(rows.CurrentPointer);
                database.Row(checked((int)(rows.CurrentPointer >> 8)), (int)(rows.CurrentPointer & 255), true, _cancellation, out uint body);
                if (!physical.Add(body) || I32(database.Page(checked((int)(body >> 8)), 1), 4) != table.DefinitionPage)
                    throw new InvalidDataException("Native application rows have aliased or foreign overflow payloads.");
            }
            foreach (AccessNativeIndex native in table.Indexes.GroupBy(x => x.RootPage).Select(x => x.First())) {
                _cancellation.ThrowIfCancellationRequested();
                if (native.Descending.Any(x => x)) throw new NotSupportedException("Native application row mutation does not qualify descending index keys.");
                if (native.Columns.Any(x => x.Type == 10 && (x.SortOrder != 1033 || x.SortVersion != 0)))
                    throw new NotSupportedException("Native application row mutation requires the qualified General legacy text collation.");
                var index = new Index(native.Name, native.Columns.Select(c => table.Columns.IndexOf(c)).ToArray(), native.Flags) { RootPage = native.RootPage };
                List<byte[]> original = ReadIndexEntries(database, table, native);
                var indexed = new HashSet<uint>();
                foreach (byte[] entry in original) if (!live.Contains(IndexRowId(entry)) || !indexed.Add(IndexRowId(entry)))
                    throw new InvalidDataException("A native edited index repeats a row or refers outside its live table records.");
                if ((native.Flags & 2) == 0 && indexed.Count != live.Count)
                    throw new InvalidDataException("A native edited index omits live table records.");
                var entries = original.Where(x => !removals.Contains(IndexRowId(x)) && (updates == null || !updates.ContainsKey(IndexRowId(x)))).ToList();
                var unique = new HashSet<string>(entries.Select(x => Convert.ToBase64String(x, 0, x.Length - 4)), StringComparer.Ordinal);
                int addedUnique = 0;
                var written = additions.Select((values, i) => (Id: rowIds[i], Values: values))
                    .Concat(updates == null ? Enumerable.Empty<(uint Id, object?[] Values)>() : updates.Select(x => (Id: x.Key, Values: x.Value)));
                var oldKeys = new HashSet<string>(original.Select(x => Convert.ToBase64String(x, 0, x.Length - 4)), StringComparer.Ordinal);
                foreach (var row in written) {
                    object?[] values = row.Values; bool nullKey = index.Columns.All(c => values[c] == null);
                    if ((index.Flags & 8) != 0 && index.Columns.Any(c => values[c] == null)) throw new InvalidDataException("A native required index cannot contain null fields.");
                    if ((index.Flags & 2) != 0 && nullKey) continue;
                    byte[] key = Key(shape, index, values, bytes => ReserveMutationMetadata(database, bytes));
                    if (unique.Add(Convert.ToBase64String(key))) { if (!oldKeys.Contains(Convert.ToBase64String(key))) addedUnique++; }
                    else if ((index.Flags & 1) != 0 && !nullKey) throw new InvalidDataException("Native application rows violate a unique index.");
                    ReserveMutationMetadata(database, key.Length + 4);
                    byte[] entry = new byte[key.Length + 4]; Buffer.BlockCopy(key, 0, entry, 0, key.Length);
                    BigEndian(entry, key.Length, row.Id); entries.Add(entry);
                }
                entries.Sort(CompareBytes); WriteIndexEntries(shape, index, entries, bytes => ReserveMutationMetadata(database, bytes));
                byte[] owned = UsageMap(database.OwnedPages(native.OwnedPages, _cancellation).Concat(index.Pages).Distinct());
                int mapPage = Allocate(); _pages[mapPage] = DataPage(0, new[] { owned });
                ReplaceDefinitionPointer(database, table.DefinitionPage, native.PhysicalDefinitionOffset + 34, checked((uint)mapPage << 8));
                // The native engine's unique-entry counter is a lifetime counter; deletions do not decrement it.
                uint count = ReadDefinitionUInt32(database, table.DefinitionPage, native.UniqueCountOffset);
                ReplaceDefinitionPointer(database, table.DefinitionPage, native.UniqueCountOffset, checked(count + (uint)addedUnique));
            }
        }

        private List<byte[]> ReadIndexEntries(AccessNativeDatabase database, AccessNativeTable table, AccessNativeIndex index) {
            var entries = new List<byte[]>(); var seen = new HashSet<int>(); var pending = new Stack<int>();
            pending.Push(index.RootPage);
            var owned = new HashSet<int>(database.OwnedPages(index.OwnedPages, _cancellation));
            while (pending.Count != 0) {
                _cancellation.ThrowIfCancellationRequested(); int number = pending.Pop();
                if (seen.Count >= database.MaxChainLength || !seen.Add(number) || !owned.Contains(number))
                    throw new InvalidDataException("A native edited index has an invalid page graph.");
                var page = database.Page(number);
                if ((page[0] != 3 && page[0] != 4) || I32(page, 4) != table.DefinitionPage)
                    throw new InvalidDataException("A native edited index has an invalid page owner or type.");
                int prefixLength = AccessNativeBinary.U16(page, 24), previous = 0;
                byte[]? prefix = null; int end = PageSize - AccessNativeBinary.U16(page, 2);
                if (end < 480 || end > PageSize) throw new InvalidDataException("A native index has invalid free-space accounting.");
                for (int offset = 0; offset < 453; offset++) for (int bit = 0; bit < 8; bit++) {
                    if ((page[27 + offset] & (1 << bit)) == 0) continue;
                    int boundary = offset * 8 + bit, length = boundary - previous;
                    if (length <= 0 || 480 + boundary > end || prefix == null && prefixLength > length)
                        throw new InvalidDataException("A native index entry overlaps its page boundary.");
                    int expanded = checked(length + (prefix?.Length ?? 0));
                    if (expanded < (page[0] == 4 ? 4 : 8) || expanded > 518)
                        throw new InvalidDataException("A native index entry exceeds the qualified metadata limit.");
                    ReserveMutationMetadata(database, expanded); byte[] entry = new byte[expanded];
                    if (prefix != null) Buffer.BlockCopy(prefix, 0, entry, 0, prefix.Length);
                    for (int i = 0; i < length; i++) entry[(prefix?.Length ?? 0) + i] = page[480 + previous + i];
                    if (prefix == null && prefixLength != 0) { ReserveMutationMetadata(database, prefixLength); prefix = new byte[prefixLength]; Buffer.BlockCopy(entry, 0, prefix, 0, prefix.Length); }
                    if (page[0] == 4) entries.Add(entry); else pending.Push(checked((int)ReadBigEndian(entry, entry.Length - 4)));
                    previous = boundary;
                }
                if (480 + previous != end) throw new InvalidDataException("A native index entry mask disagrees with its data length.");
                if (page[0] == 3) pending.Push(I32(page, 20));
            }
            entries.Sort(CompareBytes);
            for (int i = 1; i < entries.Count; i++) if (CompareBytes(entries[i - 1], entries[i]) >= 0)
                throw new InvalidDataException("A native index contains duplicate entries.");
            return entries;
        }

        private static uint ReadBigEndian(byte[] entry, int offset) => ((uint)entry[offset] << 24) | ((uint)entry[offset + 1] << 16) | ((uint)entry[offset + 2] << 8) | entry[offset + 3];
        private static uint IndexRowId(byte[] entry) => ReadBigEndian(entry, entry.Length - 4);
        private uint ReadDefinitionUInt32(AccessNativeDatabase database, int page, int offset) {
            // All counters currently edited live in the first definition page.
            if (offset < 0 || offset > PageSize - 4) throw new NotSupportedException("The native counter requires an unqualified continuation layout.");
            return AccessNativeBinary.U32(database.Page(page, 2), offset);
        }
    }
}
