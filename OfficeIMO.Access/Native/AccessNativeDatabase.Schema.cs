using OfficeIMO.Drawing;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        internal AccessNativeTable Definition(int pageNumber, string name, CancellationToken cancellation) {
            if (_definitions.TryGetValue(pageNumber, out AccessNativeTable? existing)) {
                if (!StringComparer.OrdinalIgnoreCase.Equals(existing.Name, name)) throw new InvalidDataException("Native Access catalog aliases a table definition under ambiguous names.");
                return existing;
            }
            cancellation.ThrowIfCancellationRequested();
            OfficeByteView data = DefinitionBytes(pageNumber, cancellation); int columns = U16(data, 45), logicalCount = I32(data, 47), physicalCount = I32(data, 51);
            if (columns > 255 || logicalCount < 0 || logicalCount > 256 || physicalCount < 0 || physicalCount > logicalCount) throw new InvalidDataException("Native Access table definition counts are invalid.");
            AccessNativeTable table = new AccessNativeTable(this, pageNumber, name) { RowCount = U32(data, 16), MaxColumns = U16(data, 41), MaxVariableColumns = U16(data, 43), OwnedPages = U32(data, 55) };
            if (table.MaxColumns < columns || table.MaxColumns > 8192 || table.MaxVariableColumns > table.MaxColumns) throw new InvalidDataException("Native Access table column allocation is invalid.");
            int columnStart = checked(63 + physicalCount * 12); int position = checked(columnStart + columns * 25);
            for (int i = 0; i < columns; i++) {
                cancellation.ThrowIfCancellationRequested(); int start = checked(columnStart + i * 25); OfficeByteView block = Slice(data, start, 25);
                AccessNativeColumn column = new AccessNativeColumn { Name = Name(data, ref position), Type = block[0], Number = U16(block, 5), VariableIndex = U16(block, 7),
                    Flags = block[15], ExtraFlags = block[16], FixedOffset = U16(block, 21), Size = U16(block, 23), Precision = block[11], Scale = block[12],
                    ComplexId = _document.Format == AccessFileFormat.Accdb && block[0] == 18 ? I32(block, 11) : 0 };
                if (column.Number >= table.MaxColumns || column.Variable && column.VariableIndex >= table.MaxVariableColumns) throw new InvalidDataException("Native Access field coordinates exceed the declared row schema.");
                if (table.Columns.Any(x => x.Number == column.Number || StringComparer.OrdinalIgnoreCase.Equals(x.Name, column.Name))) throw new InvalidDataException("Native Access fields have duplicate numbers or ambiguous names.");
                table.Columns.Add(column);
            }
            table.Columns.Sort((a, b) => a.Number.CompareTo(b.Number));
            List<AccessNativeIndex> physical = new List<AccessNativeIndex>();
            for (int i = 0; i < physicalCount; i++) {
                cancellation.ThrowIfCancellationRequested(); OfficeByteView block = Slice(data, position, 52); position = checked(position + 52);
                List<AccessNativeColumn> fields = new List<AccessNativeColumn>(); List<bool> descending = new List<bool>();
                for (int j = 0; j < 10; j++) {
                    int number = I16(block, 4 + j * 3); if (number == -1) continue;
                    AccessNativeColumn column = table.Columns.SingleOrDefault(x => x.Number == number) ?? throw new InvalidDataException("Native Access index refers to an unavailable field.");
                    if (fields.Contains(column)) throw new InvalidDataException("Native Access index repeats a field.");
                    fields.Add(column); descending.Add((block[6 + j * 3] & 1) == 0);
                }
                int rootPage = I32(block, 38); byte pageType = Page(rootPage)[0];
                if (pageType != 3 && pageType != 4) throw new InvalidDataException("Native Access index root has an invalid page type.");
                physical.Add(new AccessNativeIndex { Columns = fields.ToArray(), Descending = descending.ToArray(), RootPage = rootPage, Flags = block[46] });
            }
            for (int i = 0; i < logicalCount; i++) {
                OfficeByteView block = Slice(data, position, 28); position = checked(position + 28); int physicalNumber = I32(block, 8);
                if (physicalNumber < 0 || physicalNumber >= physical.Count) throw new InvalidDataException("Native Access logical index refers to an unavailable physical index.");
                AccessNativeIndex definition = physical[physicalNumber];
                table.Indexes.Add(new AccessNativeIndex { Columns = definition.Columns, Descending = definition.Descending, Flags = definition.Flags, RootPage = definition.RootPage,
                    Number = I32(block, 4), Type = block[23], RelatedIndex = I32(block, 13), RelatedTable = I32(block, 17), CascadeUpdates = (block[21] & 1) != 0, CascadeDeletes = (block[22] & 1) != 0 });
            }
            foreach (AccessNativeIndex index in table.Indexes) index.Name = Name(data, ref position);
            if (table.Indexes.Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != table.Indexes.Count) throw new InvalidDataException("Native Access index names are ambiguous.");
            _definitions.Add(pageNumber, table); return table;
        }
        private OfficeByteView DefinitionBytes(int pageNumber, CancellationToken cancellation) {
            OfficeByteView first = Page(pageNumber, 2); int length = I32(first, 8);
            if (length < 63 || length > MaxMetadataBytes) throw new InvalidDataException("Native Access table definition exceeds its metadata limit or is truncated.");
            AccountMetadata(length);
            byte[] bytes = new byte[length]; int copied = Math.Min(4096, length);
            for (int i = 0; i < copied; i++) bytes[i] = first[i];
            int next = I32(first, 4); HashSet<int> visited = new HashSet<int> { pageNumber };
            for (int depth = 1; copied < length; depth++) {
                cancellation.ThrowIfCancellationRequested();
                if (next == 0 || depth >= MaxChainLength || !visited.Add(next)) throw new InvalidDataException("Native Access table-definition chain is truncated, cyclic or exceeds its limit.");
                OfficeByteView page = Page(next, 2); int count = Math.Min(4088, length - copied);
                for (int i = 0; i < count; i++) bytes[copied + i] = page[i + 8];
                copied += count; next = I32(page, 4);
            }
            if (next != 0) throw new InvalidDataException("Native Access table-definition length disagrees with its continuation chain.");
            return bytes;
        }
        internal static AccessDataType DataType(AccessNativeColumn column) {
            if (column.Calculated) return AccessDataType.Unknown;
            switch (column.Type) {
                case 1: return AccessDataType.Boolean;
                case 2: return AccessDataType.Byte;
                case 3: return AccessDataType.Int16;
                case 4: return (column.Flags & 4) != 0 ? AccessDataType.AutoNumber : AccessDataType.Int32;
                case 5: return AccessDataType.Currency;
                case 6: return AccessDataType.Single;
                case 7: return AccessDataType.Double;
                case 8: return AccessDataType.DateTime;
                case 9: case 11: case 17: return AccessDataType.Binary;
                case 10: return AccessDataType.ShortText;
                case 12: return AccessDataType.LongText;
                case 15: return AccessDataType.Guid;
                case 19: return AccessDataType.Int64;
                case 16: return AccessDataType.Decimal;
                case 18: return AccessDataType.Complex;
                case 20: return AccessDataType.ExtendedDateTime;
                default: return AccessDataType.Unknown;
            }
        }
    }
}
