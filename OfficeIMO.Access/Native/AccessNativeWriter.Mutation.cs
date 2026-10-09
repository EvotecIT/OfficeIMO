using OfficeIMO.Drawing;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        private readonly Dictionary<AccessNativeColumn, List<int>> _mutationLongPages = new Dictionary<AccessNativeColumn, List<int>>();
        private readonly Dictionary<AccessNativeTable, List<int>> _mutationRowPages = new Dictionary<AccessNativeTable, List<int>>();
        private long _mutationMetadataBytes;
        // One writer covers every affected application/catalog/permission index.
        // Failed plans are discarded without consuming the source read budget.
        private void ReserveMutationMetadata(AccessNativeDatabase database, int bytes) {
            if (bytes < 0 || _mutationMetadataBytes > database.RemainingMetadataBytes - bytes)
                throw new InvalidDataException("Native mutation index entries exceed the remaining aggregate metadata budget.");
            _mutationMetadataBytes += bytes;
        }
        /// <summary>Plans bounded replacement of nonindexed long-binary values without changing row or catalog identities.</summary>
        /// <remarks>The original snapshot is untouched. Old payload pages remain allocated; this operation is not database compaction.</remarks>
        internal static AccessNativeWriter BuildBinaryReplacements(AccessNativeDatabase database, AccessNativeTable table,
            AccessNativeColumn column, IReadOnlyDictionary<uint, byte[]> replacements, long maximumBytes, CancellationToken cancellation) {
            if (!database.CanDecode(out _) || database.Document.Profile != AccessFormatProfile.Jet4 && database.Document.Profile != AccessFormatProfile.Ace12 && database.Document.Profile != AccessFormatProfile.Ace14)
                throw new NotSupportedException("Native application-storage mutation is limited to unprotected Jet4 and ACE12/14.");
            if (column.Type != 9 && column.Type != 11 && column.Type != 17 || column.Type == 11 && !column.Variable || column.Calculated || table.Indexes.Any(x => x.Columns.Contains(column)))
                throw new NotSupportedException("Native payload replacement requires a nonindexed binary column.");
            if (column.Type == 11 && column.OwnedLongPages == 0) throw new NotSupportedException("The native binary column has no qualified long-value usage map.");
            cancellation.ThrowIfCancellationRequested();
            byte[] snapshot = database.Snapshot();
            if (snapshot.Length > maximumBytes) throw new InvalidDataException("Native Access output exceeds MaxOutputBytes.");
            AccessNativeWriter writer = new AccessNativeWriter(maximumBytes, cancellation);
            for (int position = 0; position < snapshot.Length; position += PageSize) {
                cancellation.ThrowIfCancellationRequested(); byte[] page = new byte[PageSize];
                Buffer.BlockCopy(snapshot, position, page, 0, PageSize); writer._pages.Add(page);
            }
            if (replacements.Count == 0) return writer;
            Column payloadColumn = new Column(column.Name, 11, column.Size, variable: true);
            HashSet<uint> physicalRows = new HashSet<uint>();
            HashSet<int> tablePages = new HashSet<int>(database.OwnedPages(table.OwnedPages, cancellation));
            foreach (KeyValuePair<uint, byte[]> replacement in replacements) {
                cancellation.ThrowIfCancellationRequested();
                int rootPage = checked((int)(replacement.Key >> 8)), rootSlot = (int)(replacement.Key & 255);
                if (!tablePages.Contains(rootPage) || AccessNativeBinary.I32(database.Page(rootPage, 1), 4) != table.DefinitionPage)
                    throw new InvalidDataException("Native payload replacement refers outside its table.");
                OfficeByteView root = database.Page(rootPage);
                if (rootSlot >= AccessNativeBinary.U16(root, 12) || AccessNativeBinary.U16(root, 12) > 255)
                    throw new InvalidDataException("Native payload replacement refers outside its row directory.");
                if ((AccessNativeBinary.U16(root, 14 + rootSlot * 2) & 0x8000) != 0)
                    throw new InvalidDataException("Native payload replacement refers to a deleted row.");
                OfficeByteView row = database.Row(rootPage, rootSlot, true, cancellation, out uint physical);
                if (!tablePages.Contains(checked((int)(physical >> 8))) || AccessNativeBinary.I32(database.Page(checked((int)(physical >> 8)), 1), 4) != table.DefinitionPage)
                    throw new InvalidDataException("A native overflow payload refers outside its table.");
                if (!physicalRows.Add(physical)) throw new InvalidDataException("Native payload replacements alias the same physical row.");
                byte[] value = replacement.Value ?? throw new ArgumentException("Native replacement values cannot be null.", nameof(replacements));
                if (value.Length > database.MaxMetadataBytes) throw new InvalidDataException("Native application payload exceeds MaxMetadataBytes.");
                AccessNativeRow source = new AccessNativeRow(table, row, database.MaxMetadataBytes, metadata: true);
                byte[] descriptor = column.Type == 11 ? writer.LongValue(payloadColumn, value, external: true) : value;
                writer.ReplaceDataRow(table, checked((int)(physical >> 8)), (int)(physical & 255), source.ReplacePayload(column, descriptor));
            }
            if (column.Type == 11) {
                writer._mutationLongPages.Add(column, payloadColumn.LongPages);
                writer.UpdateLongMap(database, table, column);
            }
            if (writer._mutationRowPages.ContainsKey(table)) writer.UpdateMutationTableMaps(database, table);
            writer.WriteMutationGlobalMaps(database);
            return writer;
        }

        private void UpdateLongMap(AccessNativeDatabase database, AccessNativeTable table, AccessNativeColumn column) {
            if (column.OwnedLongPages == 0 || column.FreeLongPages == 0 || column.AmbiguousLongMaps
                || column.OwnedLongPages >> 8 != column.FreeLongPages >> 8 || column.OwnedLongPages == column.FreeLongPages)
                throw new NotSupportedException("The native long-value column lacks its owned/free map pair.");
            int[] owned = database.OwnedPages(column.OwnedLongPages, _cancellation)
                .Concat(_mutationLongPages[column]).Distinct().ToArray();
            byte[] ownedMap = UsageMap(owned);
            byte[] freeMap = database.Row(checked((int)(column.FreeLongPages >> 8)), (int)(column.FreeLongPages & 255), false, _cancellation).ToArray();
            int page = Allocate(); _pages[page] = DataPage(0, new[] { ownedMap, freeMap });
            ReplaceDefinitionPointer(database, table.DefinitionPage, column.LongMapDefinitionOffset, checked((uint)page << 8));
            ReplaceDefinitionPointer(database, table.DefinitionPage, column.LongMapDefinitionOffset + 4, checked(((uint)page << 8) | 1));
        }

        private void ReplaceDataRow(AccessNativeTable table, int number, int slot, byte[] replacement) {
            byte[] original = _pages[number]; int count = AccessNativeBinary.U16(original, 12);
            if (count > 255 || slot >= count) throw new InvalidDataException("Native replacement row is outside its page directory.");
            List<byte[]> rows = new List<byte[]>(); List<int> flags = new List<int>();
            int occupied = 14 + count * 2;
            for (int i = 0; i < count; i++) {
                _cancellation.ThrowIfCancellationRequested(); int entry = AccessNativeBinary.U16(original, 14 + i * 2);
                int start = entry & 0x1fff, end = i == 0 ? PageSize : AccessNativeBinary.U16(original, 12 + i * 2) & 0x1fff;
                if (start < 14 + count * 2 || end < start) throw new InvalidDataException("Native replacement page has invalid row offsets.");
                byte[] row = i == slot ? replacement : new byte[end - start];
                if (i != slot) Buffer.BlockCopy(original, start, row, 0, row.Length);
                occupied = checked(occupied + row.Length); rows.Add(row); flags.Add(entry & 0xe000);
            }
            if (occupied > PageSize) {
                if (replacement.Length > PageSize - 16 || rows[slot].Length < 4)
                    throw new NotSupportedException("The edited native row exceeds its qualified overflow page capacity.");
                int overflow = Allocate(); _pages[overflow] = DataPage(table.DefinitionPage, new[] { replacement });
                U16(_pages[overflow], 14, (ushort)(AccessNativeBinary.U16(_pages[overflow], 14) | 0x8000));
                if (!_mutationRowPages.TryGetValue(table, out List<int>? pages)) _mutationRowPages.Add(table, pages = new List<int>());
                pages.Add(overflow);
                // Keep the indexed root/slot identity. Overflow bodies are hidden from
                // ordinary traversal and reached only through the original row pointer.
                byte[] pointer = new byte[4]; U32(pointer, 0, checked((uint)overflow << 8));
                ReplaceDataRow(table, number, slot, pointer);
                U16(_pages[number], 14 + slot * 2, (ushort)(AccessNativeBinary.U16(_pages[number], 14 + slot * 2) | 0x4000));
                return;
            }
            byte[] output = new byte[PageSize]; Buffer.BlockCopy(original, 0, output, 0, 14);
            int offset = PageSize;
            for (int i = 0; i < count; i++) {
                offset -= rows[i].Length; Buffer.BlockCopy(rows[i], 0, output, offset, rows[i].Length);
                U16(output, 14 + i * 2, checked((ushort)(offset | flags[i])));
            }
            U16(output, 2, checked((ushort)(PageSize - occupied))); _pages[number] = output;
        }

        private void ReplaceDefinitionPointer(AccessNativeDatabase database, int first, int position, uint value) {
            int number = first, offset = position; HashSet<int> seen = new HashSet<int> { first };
            while (offset >= PageSize) {
                _cancellation.ThrowIfCancellationRequested(); offset = checked(offset - PageSize + 8);
                number = AccessNativeBinary.I32(database.Page(number, 2), 4);
                if (number == 0 || seen.Count >= database.MaxChainLength || !seen.Add(number))
                    throw new InvalidDataException("Native edited definition has an invalid continuation chain.");
            }
            for (int i = 0; i < 4; i++) {
                if (offset == PageSize) {
                    number = AccessNativeBinary.I32(database.Page(number, 2), 4); offset = 8;
                    if (number == 0 || seen.Count >= database.MaxChainLength || !seen.Add(number))
                        throw new InvalidDataException("Native edited definition has an invalid continuation chain.");
                }
                _pages[number][offset++] = (byte)(value >> (i * 8));
            }
        }
    }
}
