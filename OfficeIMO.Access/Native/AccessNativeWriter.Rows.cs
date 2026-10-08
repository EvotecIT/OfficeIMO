using System.Text;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        private void WriteRows(Table table) {
            List<byte[]> current = new List<byte[]>(); int page = Allocate(), occupied = 14;
            table.DataPages.Add(page);
            foreach (object?[] values in table.Rows) {
                _cancellation.ThrowIfCancellationRequested(); byte[] row = Row(table.Columns, values);
                if (row.Length > 4060) throw new NotSupportedException("The encoded native row exceeds the qualified capacity of 4060 bytes. Use LongText or Binary for large values.");
                if (current.Count == 255 || occupied + row.Length + 2 > PageSize) {
                    _pages[page] = DataPage(table.DefinitionPage, current); current.Clear(); occupied = 14;
                    page = Allocate(); table.DataPages.Add(page);
                }
                table.RowIds.Add(checked((uint)((page << 8) | current.Count)));
                current.Add(row); occupied += row.Length + 2;
            }
            _pages[page] = DataPage(table.DefinitionPage, current);
        }
        private byte[] Row(Column[] columns, object?[] values) {
            using MemoryStream data = new MemoryStream(); using BinaryWriter writer = new BinaryWriter(data, Encoding.Unicode, true);
            writer.Write((short)columns.Length); byte[] mask = new byte[(columns.Length + 7) / 8];
            for (int i = 0; i < columns.Length; i++) {
                Column column = columns[i]; object? value = values[i];
                if (value != null && (column.Type != 1 || (bool)value)) mask[i / 8] |= (byte)(1 << (i % 8));
                if (column.Variable) continue;
                switch (column.Type) {
                    case 1: break;
                    case 2: writer.Write(value == null ? (byte)0 : (byte)value); break;
                    case 3: writer.Write(value == null ? (short)0 : (short)value); break;
                    case 4: writer.Write(value == null ? 0 : (int)value); break;
                    case 5: writer.Write(value == null ? 0L : checked((long)((decimal)value * 10000m))); break;
                    case 6: writer.Write(value == null ? 0f : (float)value); break;
                    case 7: writer.Write(value == null ? 0d : (double)value); break;
                    case 8:
                        double native = value is DateTime date ? date.ToOADate() : value == null ? 0d : (double)value;
                        if (value is DateTime expected && DateTime.FromOADate(native).Ticks != expected.Ticks) throw new NotSupportedException("DateTime cannot be represented exactly by this native OLE Automation date codec.");
                        writer.Write(native); break;
                    case 15: writer.Write(value == null ? new byte[16] : ((Guid)value).ToByteArray()); break;
                    case 16: writer.Write(Numeric(column, value == null ? 0m : (decimal)value)); break;
                    default: throw new NotSupportedException("The native fixed field type is unqualified for creation.");
                }
            }
            List<ushort> offsets = new List<ushort>();
            for (int i = 0; i < columns.Length; i++) {
                Column column = columns[i]; if (!column.Variable) continue;
                offsets.Add(checked((ushort)data.Position)); object? value = values[i];
                if (value == null) continue;
                byte[] bytes;
                if (value is string text) {
                    int count = Encoding.Unicode.GetByteCount(text);
                    if (count > _maximumBytes) throw new InvalidDataException("A text value exceeds MaxOutputBytes.");
                    bytes = new UnicodeEncoding(false, false, true).GetBytes(text);
                } else bytes = (byte[])value;
                if (bytes.Length > _maximumBytes) throw new InvalidDataException("A binary value exceeds MaxOutputBytes.");
                writer.Write(column.Type == 11 || column.Type == 12 ? LongValue(column, bytes) : bytes);
            }
            if (offsets.Count != 0) { writer.Write(checked((ushort)data.Position)); for (int i = offsets.Count - 1; i >= 0; i--) writer.Write(offsets[i]); writer.Write((ushort)offsets.Count); }
            writer.Write(mask); writer.Flush(); return data.ToArray();
        }
        private static byte[] Numeric(Column column, decimal value) {
            if (column.Precision < 1 || column.Precision > 28 || column.Scale > column.Precision) throw new NotSupportedException("Native Decimal requires precision 1–28 and scale no greater than precision.");
            decimal factor = 1m; for (int i = 0; i < column.Scale; i++) factor *= 10m;
            decimal magnitude = checked(Math.Abs(value) * factor);
            if (decimal.Truncate(magnitude) != magnitude) throw new NotSupportedException("Decimal exceeds the declared scale; native creation never rounds it.");
            decimal maximum = 1m; for (int i = 0; i < column.Precision; i++) maximum *= 10m;
            if (magnitude >= maximum) throw new NotSupportedException("Decimal exceeds the declared precision.");
            int[] bits = decimal.GetBits(decimal.Truncate(magnitude)); byte[] output = new byte[17]; output[0] = value < 0 ? (byte)0x80 : (byte)0;
            U32(output, 5, unchecked((uint)bits[2])); U32(output, 9, unchecked((uint)bits[1])); U32(output, 13, unchecked((uint)bits[0])); return output;
        }
        private byte[] LongValue(Column column, byte[] value) {
            byte[] descriptor = new byte[12];
            if (value.Length <= 64) {
                byte[] inline = new byte[value.Length + 12]; U32(inline, 0, (uint)value.Length | 0x80000000);
                Buffer.BlockCopy(value, 0, inline, 12, value.Length); return inline;
            }
            if (value.Length <= 4076) {
                int number = Allocate(); column.LongPages.Add(number); byte[] page = DataPage(0, new[] { value });
                page[4] = (byte)'L'; page[5] = (byte)'V'; page[6] = (byte)'A'; page[7] = (byte)'L'; _pages[number] = page;
                U32(descriptor, 0, (uint)value.Length | 0x40000000); U32(descriptor, 4, (uint)(number << 8)); return descriptor;
            }
            // Microsoft Jet/ACE requires full intermediate chunks of 4072 payload bytes, plus the four-byte next-row pointer.
            const int capacity = 4072;
            int chunks = checked((value.Length + capacity - 1) / capacity);
            int[] pages = new int[chunks]; for (int i = 0; i < chunks; i++) { pages[i] = Allocate(); column.LongPages.Add(pages[i]); }
            int offset = 0;
            for (int i = 0; i < chunks; i++) {
                int length = Math.Min(capacity, value.Length - offset); byte[] row = new byte[length + 4];
                U32(row, 0, i + 1 < chunks ? (uint)(pages[i + 1] << 8) : 0);
                Buffer.BlockCopy(value, offset, row, 4, length); offset += length;
                byte[] page = DataPage(0, new[] { row }); page[4] = (byte)'L'; page[5] = (byte)'V'; page[6] = (byte)'A'; page[7] = (byte)'L'; _pages[pages[i]] = page;
            }
            U32(descriptor, 0, (uint)value.Length); U32(descriptor, 4, (uint)(pages[0] << 8)); return descriptor;
        }
        private static byte[] DataPage(int owner, IReadOnlyList<byte[]> rows) {
            byte[] page = new byte[PageSize]; page[0] = 1; page[1] = 1; U32(page, 4, (uint)owner); U16(page, 12, checked((ushort)rows.Count));
            int end = PageSize;
            for (int i = 0; i < rows.Count; i++) { end -= rows[i].Length; if (end < 14 + rows.Count * 2) throw new InvalidDataException("Native rows exceed their allocated page."); Buffer.BlockCopy(rows[i], 0, page, end, rows[i].Length); U16(page, 14 + i * 2, (ushort)end); }
            U16(page, 2, (ushort)(end - 14 - rows.Count * 2)); return page;
        }
    }
}
