using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        /// <summary>Appends native rows and marks removed root rows without moving any surviving row identity.</summary>
        internal void MutateApplicationRows(AccessNativeDatabase database, AccessNativeTable table,
            IReadOnlyList<object?[]> additions, ISet<uint> removals, IReadOnlyDictionary<uint, string>? renamedRows = null, long? rowLimit = null) {
            if (additions.Count == 0 && removals.Count == 0 && (renamedRows == null || renamedRows.Count == 0)) return;
            long limit = rowLimit ?? database.MaxCatalogObjects;
            if (limit < 1 || table.RowCount + additions.Count - removals.Count > limit)
                throw new InvalidDataException("Native application rows exceed their qualified row limit.");
            Column[] columns = MutationColumns(table);
            var live = new HashSet<uint>();
            var updated = new Dictionary<uint, object?[]>();
            using (var cursor = new AccessNativeRowCursor(table, _cancellation, rowLimit: limit))
                while (cursor.Read(_cancellation)) {
                    live.Add(cursor.CurrentPointer);
                    if (renamedRows == null || !renamedRows.TryGetValue(cursor.CurrentPointer, out string? name)) continue;
                    AccessNativeColumn column = table.Columns.Single(x => x.Name.Equals("Name", StringComparison.OrdinalIgnoreCase));
                    object?[] values = new object?[columns.Length];
                    foreach (AccessNativeColumn indexed in table.Indexes.SelectMany(x => x.Columns).Distinct()) {
                        int ordinal = table.Columns.IndexOf(indexed); values[ordinal] = cursor.GetValue(ordinal, _cancellation);
                    }
                    values[table.Columns.IndexOf(column)] = name; updated.Add(cursor.CurrentPointer, values);
                    database.Row(checked((int)(cursor.CurrentPointer >> 8)), (int)(cursor.CurrentPointer & 255), true, _cancellation, out uint physical);
                    if (I32(database.Page(checked((int)(physical >> 8)), 1), 4) != table.DefinitionPage)
                        throw new InvalidDataException("A native renamed row overflows outside its table.");
                    ReplaceDataRow(table, checked((int)(physical >> 8)), (int)(physical & 255), cursor.Current.ReplacePayload(column, AccessNativeBinary.Unicode.GetBytes(name)));
                }
            if (renamedRows != null && (updated.Count != renamedRows.Count || renamedRows.Keys.Any(removals.Contains)))
                throw new InvalidDataException("Native renamed application rows are missing or conflict with deletion.");
            if (removals.Any(x => !live.Contains(x))) throw new InvalidDataException("A removed native application row is not live in its table.");
            foreach (uint id in removals) {
                int page = checked((int)(id >> 8)), offset = 14 + (int)(id & 255) * 2;
                U16(_pages[page], offset, (ushort)(AccessNativeBinary.U16(_pages[page], offset) | 0x8000));
            }
            var pages = new List<int>(); var rowIds = new List<uint>();
            int current = -1, occupied = 14; var rows = new List<byte[]>();
            foreach (object?[] values in additions) {
                _cancellation.ThrowIfCancellationRequested();
                if (values.Length != columns.Length) throw new ArgumentException("Native application row values disagree with their schema.", nameof(additions));
                byte[] row = Row(columns, values);
                if (row.Length > 4060) throw new NotSupportedException("A new native application row exceeds its qualified page capacity.");
                if (current < 0 || rows.Count == 255 || occupied + row.Length + 2 > PageSize) {
                    if (current >= 0) _pages[current] = DataPage(table.DefinitionPage, rows);
                    current = Allocate(); pages.Add(current); rows.Clear(); occupied = 14;
                }
                rowIds.Add(checked(((uint)current << 8) | (uint)rows.Count)); rows.Add(row); occupied += row.Length + 2;
            }
            if (current >= 0) _pages[current] = DataPage(table.DefinitionPage, rows);
            UpdateMutationIndexes(database, table, additions, rowIds, removals, columns, updated, limit);
            if (!_mutationRowPages.TryGetValue(table, out List<int>? addedPages)) _mutationRowPages.Add(table, addedPages = new List<int>());
            addedPages.AddRange(pages); UpdateMutationTableMaps(database, table);
            ReplaceDefinitionPointer(database, table.DefinitionPage, 16, checked((uint)(table.RowCount + additions.Count - removals.Count)));
            for (int i = 0; i < columns.Length; i++) {
                if (columns[i].LongPages.Count != 0) {
                    if (!_mutationLongPages.TryGetValue(table.Columns[i], out List<int>? values)) _mutationLongPages.Add(table.Columns[i], values = new List<int>());
                    values.AddRange(columns[i].LongPages); UpdateLongMap(database, table, table.Columns[i]);
                }
                if ((table.Columns[i].Flags & 4) != 0 && additions.Count != 0) {
                    if (columns[i].Type != 4) throw new NotSupportedException("Native application creation does not qualify this AutoNumber field type.");
                    int maximum = Math.Max(I32(database.Page(table.DefinitionPage, 2), 20), additions.Max(row => Convert.ToInt32(row[i])));
                    ReplaceDefinitionPointer(database, table.DefinitionPage, 20, unchecked((uint)maximum));
                }
            }
            WriteMutationGlobalMaps(database);
        }

        /// <summary>Retains the native owned/free map pair while adding new data and overflow pages.</summary>
        private void UpdateMutationTableMaps(AccessNativeDatabase database, AccessNativeTable table) {
            List<int> added = _mutationRowPages[table];
            byte[] owned = UsageMap(database.OwnedPages(table.OwnedPages, _cancellation).Concat(added).Distinct());
            byte[] free = UsageMap(database.OwnedPages(table.FreePages, _cancellation).Concat(added).Distinct()
                .Where(p => _pages[p][0] == 1 && I32(_pages[p], 4) == table.DefinitionPage && AccessNativeBinary.U16(_pages[p], 12) < 255 && AccessNativeBinary.U16(_pages[p], 2) > 2));
            int mapPage = Allocate(); _pages[mapPage] = DataPage(0, new[] { owned, free });
            ReplaceDefinitionPointer(database, table.DefinitionPage, 55, checked((uint)mapPage << 8));
            ReplaceDefinitionPointer(database, table.DefinitionPage, 59, checked(((uint)mapPage << 8) | 1));
        }

        private static Column[] MutationColumns(AccessNativeTable table) {
            if (table.MaxColumns != table.Columns.Count || table.MaxVariableColumns != table.Columns.Count(c => c.Variable))
                throw new NotSupportedException("Native application row creation does not qualify dropped or sparse field coordinates.");
            int fixedOffset = 0, variable = 0;
            var columns = new List<Column>();
            for (int i = 0; i < table.Columns.Count; i++) {
                AccessNativeColumn source = table.Columns[i];
                if (source.Number != i || source.Calculated || source.Variable && source.VariableIndex != variable++
                    || !source.Variable && source.Type != 1 && source.FixedOffset != fixedOffset)
                    throw new NotSupportedException("The native application row schema has unqualified field coordinates.");
                if (!source.Variable && source.Type != 1) fixedOffset += source.Size;
                columns.Add(new Column(source.Name, source.Type, source.Size, source.Variable) { Precision = source.Precision, Scale = source.Scale });
            }
            return columns.ToArray();
        }
    }
}
