using System.Text;
namespace OfficeIMO.Access;
internal sealed partial class AccessNativeWriter {
    private byte[] DefinitionBody(Table table) {
        using var content = new MemoryStream(); using var writer = new BinaryWriter(content, AccessNativeBinary.Unicode, true);
        writer.Write(0); writer.Write(1625); writer.Write(table.Rows.Length); writer.Write(table.AutoNumberLast); writer.Write(1);
        writer.Write(0); writer.Write(0L); writer.Write((byte)(table.Name.StartsWith("MSys", StringComparison.Ordinal) ? 0x53 : 0x4e));
        writer.Write((short)table.Columns.Length); writer.Write((short)table.Columns.Count(c => c.Variable)); writer.Write((short)table.Columns.Length);
        Index[] physicalIndexes = PhysicalIndexes(table);
        writer.Write(table.Indexes.Count); writer.Write(physicalIndexes.Length); writer.Write(table.MapPointers[0]); writer.Write(table.MapPointers[1]);
        foreach (Index index in physicalIndexes) { writer.Write(0); writer.Write(index.UniqueCount); writer.Write(0); }
        var blocks = new List<(Column Column, int Number, int Variable, int Fixed)>();
        int variable = 0, fixedOffset = 0;
        for (int i = 0; i < table.Columns.Length; i++) {
            Column column = table.Columns[i];
            blocks.Add((column, i, variable, fixedOffset));
            if (column.Variable) variable++; else if (column.Type != 1) fixedOffset += column.Size;
        }
        foreach (var block in blocks) {
            Column column = block.Column;
            writer.Write(column.Type); writer.Write(1625); writer.Write((short)block.Number);
            writer.Write((short)block.Variable); writer.Write((short)block.Number);
            writer.Write(column.Type == 16 ? column.Precision | (column.Scale << 8) : column.Type == 10 || column.Type == 12 ? 1033 : 0);
            byte flags = (byte)(column.Variable ? 2 : 3);
            if (column.IsSystemSid) flags |= 0x30;
            if (table.Name.StartsWith("MSys", StringComparison.Ordinal)) flags |= 0x10;
            if (column.AutoNumber) flags |= 4;
            writer.Write(flags); writer.Write((byte)0); writer.Write(0);
            writer.Write((short)(column.Variable || column.Type == 1 ? 0 : block.Fixed)); writer.Write((short)column.Size);
        }
        foreach (var block in blocks) { byte[] name = AccessNativeBinary.Unicode.GetBytes(block.Column.Name); writer.Write((short)name.Length); writer.Write(name); }
        for (int i = 0; i < physicalIndexes.Length; i++) {
            Index index = physicalIndexes[i]; writer.Write(1923);
            for (int j = 0; j < 10; j++) { writer.Write((short)(j < index.Columns.Length ? index.Columns[j] : -1)); writer.Write((byte)(j < index.Columns.Length ? 1 : 0)); }
            writer.Write(table.MapPointers[i + 2]); writer.Write(index.RootPage); writer.Write(0); writer.Write(index.Flags); writer.Write((byte)0); writer.Write(0);
        }
        foreach (var indexed in table.Indexes.Select((index, i) => (index, i)).OrderBy(x => x.index.Name, StringComparer.OrdinalIgnoreCase)) {
            Index index = indexed.index; int physical = Array.FindIndex(physicalIndexes, candidate => candidate.RootPage == index.RootPage);
            writer.Write(1625); writer.Write(indexed.i); writer.Write(physical); writer.Write(index.RelatedType); writer.Write(index.RelatedIndex); writer.Write(index.RelatedTable); writer.Write((byte)0); writer.Write((byte)0); writer.Write(index.Type); writer.Write(0);
        }
        foreach (Index index in table.Indexes.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)) { byte[] name = AccessNativeBinary.Unicode.GetBytes(index.Name); writer.Write((short)name.Length); writer.Write(name); }
        int longMapRow = 2 + physicalIndexes.Length;
        for (int i = 0; i < table.Columns.Length; i++) {
            if (table.Columns[i].Type != 11 && table.Columns[i].Type != 12) continue;
            writer.Write((short)i); writer.Write(table.MapPointers[longMapRow++]); writer.Write(table.MapPointers[longMapRow++]);
        }
        writer.Write((short)-1); writer.Flush();
        // The declared definition length includes the page header; its free-space accounting reserves a further eight bytes.
        byte[] body = content.ToArray(); U32(body, 0, (uint)(body.Length + 8));
        return body;
    }
    private void WriteDefinition(Table table) {
        byte[] body = DefinitionBody(table);
        int offset = 0, pageNumber = table.DefinitionPage;
        while (offset < body.Length) {
            _cancellation.ThrowIfCancellationRequested();
            byte[] page = _pages[pageNumber]; page[0] = 2; page[1] = 1;
            int count = Math.Min(PageSize - 8, body.Length - offset);
            Buffer.BlockCopy(body, offset, page, 8, count); offset += count;
            U16(page, 2, (ushort)Math.Max(0, PageSize - 16 - count));
            if (offset < body.Length) { int next = Allocate(); U32(page, 4, (uint)next); pageNumber = next; }
        }
    }
}
