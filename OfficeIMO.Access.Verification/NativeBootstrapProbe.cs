using System.Buffers.Binary;
using System.Text;

namespace OfficeIMO.Access.Verification;

/// <summary>Opt-in fixed native feasibility probe. Independent DAO qualification is separate from the production Save codec.</summary>
internal static class NativeBootstrapProbe {
    private const int PageSize = 4096;
    private sealed record Column(string Name, byte Type, int Size, bool Variable = false);
    private sealed record Index(string Name, int RootPage, int[] Columns, byte Flags, byte Type = 0, int RelatedTable = 0, int RelatedIndex = -1, byte RelatedType = 0);
    private sealed record Table(int DefinitionPage, int MapPage, int DataPage, string Name, Column[] Columns, object?[][] Rows, Index[] Indexes);

    internal static void Generate(string outputDirectory) {
        string root = Path.GetFullPath(outputDirectory);
        if (Directory.Exists(root) || File.Exists(root)) throw new IOException("Bootstrap probes require a fresh directory.");
        Directory.CreateDirectory(root);
        foreach (bool ace in new[] { false, true }) {
            var catalogColumns = new[] {
                new Column("Id", 4, 4), new Column("ParentId", 4, 4), new Column("Name", 10, 510, true),
                new Column("Type", 3, 2), new Column("DateCreate", 8, 8), new Column("DateUpdate", 8, 8),
                new Column("Owner", 9, 255, true), new Column("Flags", 4, 4),
                new Column("Database", 12, 0, true), new Column("Connect", 12, 0, true), new Column("ForeignName", 10, 510, true),
                new Column("RmtInfoShort", 9, 255, true), new Column("RmtInfoLong", 11, 0, true), new Column("Lv", 11, 0, true),
                new Column("LvProp", 11, 0, true), new Column("LvModule", 11, 0, true), new Column("LvExtra", 11, 0, true)
            };
            object?[] CatalogRow(int id, int parent, string name, short type) =>
                new object?[] { id, parent, name, type, 0d, 0d, new byte[] { 3, 1 }, int.MinValue, null, null, null, null, null, null, null, null, null };
            var catalogRows = new[] {
                CatalogRow(0x0f000001, 0x0f000000, "Tables", 3),
                CatalogRow(0x0f000002, 0x0f000000, "Databases", 3),
                CatalogRow(0x0f000003, 0x0f000000, "Relationships", 3),
                CatalogRow(0x10000000, 0x0f000002, "MSysDb", 2),
                CatalogRow(2, 0x0f000001, "MSysObjects", 1), CatalogRow(3, 0x0f000001, "MSysACEs", 1),
                CatalogRow(4, 0x0f000001, "MSysQueries", 1), CatalogRow(5, 0x0f000001, "MSysRelationships", 1),
                CatalogRow(21,0x0f000001,"Groups",1),CatalogRow(25,0x0f000001,"Contacts",1),CatalogRow(int.MinValue,0x0f000003,"ContactGroups",8)
            };
            foreach (object?[] row in catalogRows.Skip(8)) row[7] = 0;
            var tables = new[] {
                new Table(2, 6, 7, "MSysObjects", catalogColumns, catalogRows, new[] { new Index("ParentIdName",14,new[]{1,2},129), new Index("Id",15,new[]{0},129,1) }),
                new Table(3, 8, 9, "MSysACEs", new[] { new Column("ObjectId",4,4), new Column("SID",9,510,true), new Column("ACM",4,4), new Column("FInheritable",1,1) }, PermissionRows(), new[] {new Index("ObjectId",16,new[]{0},136)}),
                new Table(4, 10, 11, "MSysQueries", new[] { new Column("ObjectId",4,4), new Column("Attribute",2,1), new Column("Order",9,510,true), new Column("Name1",10,510,true), new Column("Name2",10,510,true), new Column("Expression",12,0,true), new Column("Flag",3,2), new Column("LvExtra",4,4) }, Array.Empty<object?[]>(), new[] {new Index("ObjectIdAttribute",17,new[]{0,1,2},129,1)}),
                new Table(5, 12, 13, "MSysRelationships", new[] { new Column("szRelationship",10,510,true), new Column("grbit",4,4), new Column("ccolumn",4,4), new Column("icolumn",4,4), new Column("szObject",10,510,true), new Column("szColumn",10,510,true), new Column("szReferencedObject",10,510,true), new Column("szReferencedColumn",10,510,true) }, new[]{new object?[]{"ContactGroups",0,1,0,"Contacts","GroupId","Groups","Id"}}, new[] {new Index("szRelationship",18,new[]{0},130),new Index("szObject",19,new[]{4},130),new Index("szReferencedObject",20,new[]{6},130)}),
                new Table(21,22,23,"Groups",new[]{new Column("Id",4,4),new Column("Label",10,160,true)},new[]{new object?[]{1,"Group A"}},new[]{new Index("PK_Groups",24,new[]{0},137,1),new Index(".rB",24,new[]{0},137,2,25,1,1)}),
                new Table(25,26,27,"Contacts",new[]{new Column("Id",4,4),new Column("GroupId",4,4),new Column("DisplayName",10,240,true),new Column("Amount",5,8),new Column("CreatedAt",8,8),new Column("Active",1,1)},new[]{new object?[]{1,1,"Ada",12.3456m,new DateTime(2026,1,2,3,4,5).ToOADate(),true},new object?[]{2,1,"",-1.25m,null,false}},new[]{new Index("PK_Contacts",28,new[]{0},137,1),new Index("ContactGroups",29,new[]{1},128,2,21,1,2)})
            };
            var pages = Enumerable.Range(0, 30).Select(_ => new byte[PageSize]).ToArray();
            pages[0] = Header(ace);
            byte[] sidMask = SidMask(pages[0]);
            foreach (Table table in tables) foreach (object?[] row in table.Rows) for (int i = 0; i < table.Columns.Length; i++) {
                        if ((table.Columns[i].Name == "Owner" || table.Columns[i].Name == "SID") && row[i] is byte[] sid)
                            row[i] = sid.Select((value, index) => (byte)(value ^ sidMask[index])).ToArray();
                    }
            byte[] global = UsageMap(Enumerable.Range(0, pages.Length));
            // The global map marks free pages; per-table maps mark owned pages.
            for (int i = 5; i < global.Length; i++) global[i] = (byte)~global[i];
            pages[1] = DataPage(1, new[] { global, global });
            foreach (Table table in tables) {
                pages[table.DefinitionPage] = Definition(table);
                byte[] usage = UsageMap(new[] { table.DataPage });
                int longColumns = table.Columns.Count(c => c.Type == 11 || c.Type == 12);
                pages[table.MapPage] = DataPage(0, new[] { usage, usage }.Concat(table.Indexes.DistinctBy(index => index.RootPage).Select(index => UsageMap(new[] { index.RootPage }))).Concat(Enumerable.Range(0, longColumns * 2).Select(_ => UsageMap(Array.Empty<int>()))).ToArray());
                pages[table.DataPage] = DataPage(table.DefinitionPage, table.Rows.Select(row => Row(table.Columns, row)).ToArray());
                foreach (Index index in table.Indexes.DistinctBy(index => index.RootPage)) pages[index.RootPage] = IndexPage(table, index);
            }
            using var output = new FileStream(Path.Combine(root, ace ? "catalog-scaffold.accdb" : "catalog-scaffold.mdb"), FileMode.CreateNew);
            foreach (byte[] page in pages) output.Write(page);
        }
    }

    private static object?[][] PermissionRows() {
        object?[] Permission(int id, int sid, int rights, bool inherited = false) => new object?[] { id, new byte[] { (byte)sid, (byte)(sid >> 8) }, rights, inherited };
        return new[] {
            // Plain built-in IDs observed by decoding two independently produced synthetic databases.
            Permission(2,0x0103,393216),Permission(3,0x0103,393216),Permission(4,0x0103,393216),Permission(5,0x0103,917504),
            Permission(0x0f000001,0x0402,983294,true),Permission(0x0f000001,0x0103,393217),
            Permission(0x0f000003,0x0402,983294,true),Permission(0x0f000003,0x0103,393217),
            Permission(0x0f000002,0x0103,393216),Permission(0x10000000,0x0103,393230),Permission(0x10000000,0x0102,14),
            Permission(4,0x0102,20),Permission(5,0x0102,20),Permission(2,0x0102,20),
            Permission(0x0f000001,0x0102,1048319,true),Permission(0x0f000003,0x0102,1048575,true),
            Permission(21,0x0103,983294),Permission(21,0x0102,1048319),Permission(25,0x0103,983294),Permission(25,0x0102,1048319),
            Permission(int.MinValue,0x0103,983294),Permission(int.MinValue,0x0102,1048575)
        };
    }

    private static byte[] Header(bool ace) {
        var page = new byte[PageSize]; page[1] = 1;
        Encoding.ASCII.GetBytes(ace ? "Standard ACE DB" : "Standard Jet DB").CopyTo(page, 4);
        page[20] = (byte)(ace ? 2 : 1);
        // Logical fixed-root declarations and a reproducible creation date for the unprotected profile.
        U32(page, 24, 0x100); U32(page, 28, 0x101);
        for (int root = 2; root <= 5; root++) U32(page, 24 + root * 4, (uint)root);
        U16(page, 60, 1252); U32(page, 110, 1033);
        double date = new DateTime(2026, 1, 1).ToOADate();
        BinaryPrimitives.WriteInt64LittleEndian(page.AsSpan(114, 8), BitConverter.DoubleToInt64Bits(date));
        byte[] dateMask = new byte[4]; BinaryPrimitives.WriteInt32LittleEndian(dateMask, (int)date);
        for (int i = 0; i < 40; i++) page[66 + i] = dateMask[i % 4];
        // Observed unprotected header field at 0x6A; further security-profile semantics remain unqualified.
        U32(page, 106, 4518);
        U32(page, 152, 1620); Encoding.ASCII.GetBytes("4.0").CopyTo(page, 156);
        // Empty Jet/ACE user-slot declarations; zero-filled slots are interpreted as occupied by DAO.
        page[3584] = 1; for (int slot = 3585; slot < PageSize; slot += 2) page[slot] = 1;
        // Jet's fixed header obfuscation is separate from database password/encryption protection.
        byte[] pad = Rc4Pad(new byte[] { 0xc7, 0xda, 0x39, 0x6b }, 128);
        for (int n = 0; n < pad.Length; n++) page[24 + n] ^= pad[n];
        return page;
    }

    private static byte[] SidMask(byte[] maskedHeader) {
        // Jet/ACE SID fields use their own RC4 key folded from the clear header, including its date and password region.
        // Physical contract reference: Jackcess Encrypt's public SidRemasker / JetPasswordHandler documentation.
        byte[] header = (byte[])maskedHeader.Clone(); byte[] headerMask = Rc4Pad(new byte[] { 0xc7, 0xda, 0x39, 0x6b }, 128);
        for (int i = 0; i < headerMask.Length; i++) header[24 + i] ^= headerMask[i];
        uint folded = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(114, 4));
        int dateMask = (int)BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(header.AsSpan(114, 8)));
        for (int i = 0; i < 40; i++) {
            int position = i * 2; byte value = header[66 + position];
            if (position < 40) value ^= (byte)(dateMask >> ((position % 4) * 8));
            folded ^= (uint)value << (i % 24);
        }
        byte[] key = new byte[4]; BinaryPrimitives.WriteUInt32LittleEndian(key, folded); return Rc4Pad(key, 2);
    }

    private static byte[] Rc4Pad(byte[] key, int length) {
        byte[] state = Enumerable.Range(0, 256).Select(i => (byte)i).ToArray();
        int j = 0;
        for (int i = 0; i < 256; i++) { j = (j + state[i] + key[i % key.Length]) & 255; (state[i], state[j]) = (state[j], state[i]); }
        int x = 0; j = 0;
        var pad = new byte[length];
        for (int n = 0; n < length; n++) { x = (x + 1) & 255; j = (j + state[x]) & 255; (state[x], state[j]) = (state[j], state[x]); pad[n] = state[(state[x] + state[j]) & 255]; }
        return pad;
    }

    private static byte[] Definition(Table table) {
        using var content = new MemoryStream(); using var writer = new BinaryWriter(content, Encoding.Unicode, true);
        writer.Write(0); writer.Write(1625); writer.Write(table.Rows.Length); writer.Write(0); writer.Write(1);
        writer.Write(0); writer.Write(0L); writer.Write((byte)(table.Name.StartsWith("MSys", StringComparison.Ordinal) ? 0x53 : 0x4e));
        writer.Write((short)table.Columns.Length); writer.Write((short)table.Columns.Count(c => c.Variable)); writer.Write((short)table.Columns.Length);
        Index[] physicalIndexes = table.Indexes.DistinctBy(index => index.RootPage).ToArray();
        writer.Write(table.Indexes.Length); writer.Write(physicalIndexes.Length); writer.Write(table.MapPage << 8); writer.Write((table.MapPage << 8) | 1);
        foreach (Index index in physicalIndexes) { writer.Write(0); writer.Write(table.Rows.Select(row => string.Join("|", index.Columns.Select(column => row[column]))).Distinct().Count()); writer.Write(0); }
        var blocks = new List<(Column Column, int Number, int Variable, int Fixed)>();
        int variable = 0, fixedOffset = 0;
        for (int i = 0; i < table.Columns.Length; i++) {
            Column column = table.Columns[i];
            blocks.Add((column, i, variable, fixedOffset));
            if (column.Variable) variable++; else if (column.Type != 1) fixedOffset += column.Size;
        }
        foreach (var block in blocks.OrderBy(b => b.Column.Name, StringComparer.OrdinalIgnoreCase)) {
            Column column = block.Column;
            writer.Write(column.Type); writer.Write(1625); writer.Write((short)block.Number);
            writer.Write((short)block.Variable); writer.Write((short)block.Number);
            writer.Write(column.Type == 10 || column.Type == 12 ? 1033 : 0);
            byte flags = (byte)(column.Name == "Owner" || column.Name == "SID" ? 0x32 : column.Variable ? 2 : 3);
            if (table.Name.StartsWith("MSys", StringComparison.Ordinal)) flags |= 0x10;
            writer.Write(flags); writer.Write((byte)0); writer.Write(0);
            writer.Write((short)(column.Variable || column.Type == 1 ? 0 : block.Fixed)); writer.Write((short)column.Size);
        }
        foreach (var block in blocks.OrderBy(b => b.Column.Name, StringComparer.OrdinalIgnoreCase)) { byte[] name = Encoding.Unicode.GetBytes(block.Column.Name); writer.Write((short)name.Length); writer.Write(name); }
        for (int i = 0; i < physicalIndexes.Length; i++) {
            Index index = physicalIndexes[i]; writer.Write(1923);
            for (int j = 0; j < 10; j++) { writer.Write((short)(j < index.Columns.Length ? index.Columns[j] : -1)); writer.Write((byte)(j < index.Columns.Length ? 1 : 0)); }
            writer.Write((table.MapPage << 8) | (i + 2)); writer.Write(index.RootPage); writer.Write(0); writer.Write(index.Flags); writer.Write((byte)0); writer.Write(0);
        }
        foreach (var indexed in table.Indexes.Select((index, i) => (index, i)).OrderBy(x => x.index.Name, StringComparer.OrdinalIgnoreCase)) {
            Index index = indexed.index; int physical = Array.FindIndex(physicalIndexes, candidate => candidate.RootPage == index.RootPage);
            writer.Write(1625); writer.Write(indexed.i); writer.Write(physical); writer.Write(index.RelatedType); writer.Write(index.RelatedIndex); writer.Write(index.RelatedTable); writer.Write((byte)0); writer.Write((byte)0); writer.Write(index.Type); writer.Write(0);
        }
        foreach (Index index in table.Indexes.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase)) { byte[] name = Encoding.Unicode.GetBytes(index.Name); writer.Write((short)name.Length); writer.Write(name); }
        int longMapRow = 2 + physicalIndexes.Length;
        for (int i = 0; i < table.Columns.Length; i++) {
            if (table.Columns[i].Type != 11 && table.Columns[i].Type != 12) continue;
            writer.Write((short)i); writer.Write((table.MapPage << 8) | longMapRow++); writer.Write((table.MapPage << 8) | longMapRow++);
        }
        writer.Write((short)-1); writer.Flush();
        // The declared definition length includes the page header; its free-space accounting reserves a further eight bytes.
        byte[] body = content.ToArray(); U32(body, 0, (uint)(body.Length + 8));
        if (body.Length > PageSize - 8) throw new InvalidDataException("Probe table definition exceeds one page.");
        var page = new byte[PageSize]; page[0] = 2; page[1] = 1; U16(page, 2, (ushort)(PageSize - 16 - body.Length)); body.CopyTo(page, 8); return page;
    }

    private static byte[] Row(Column[] columns, object?[] values) {
        using var data = new MemoryStream(); using var writer = new BinaryWriter(data, Encoding.Unicode, true);
        writer.Write((short)columns.Length); var mask = new byte[(columns.Length + 7) / 8];
        for (int i = 0; i < columns.Length; i++) {
            if (values[i] != null && (columns[i].Type != 1 || (bool)values[i]!)) mask[i / 8] |= (byte)(1 << (i % 8));
            if (columns[i].Variable) continue;
            switch (columns[i].Type) {
                case 2: writer.Write(values[i] == null ? (byte)0 : Convert.ToByte(values[i])); break;
                case 3: writer.Write(values[i] == null ? (short)0 : Convert.ToInt16(values[i])); break;
                case 4: writer.Write(values[i] == null ? 0 : Convert.ToInt32(values[i])); break;
                case 5: writer.Write(values[i] == null ? 0L : checked((long)((decimal)values[i]! * 10000m))); break;
                case 8: writer.Write(values[i] == null ? 0d : Convert.ToDouble(values[i])); break;
                case 1: break;
                default: throw new NotSupportedException("Probe fixed type is not implemented.");
            }
        }
        var offsets = new List<short>();
        for (int i = 0; i < columns.Length; i++) {
            if (!columns[i].Variable) continue;
            offsets.Add(checked((short)data.Position));
            if (values[i] is string text) writer.Write(Encoding.Unicode.GetBytes(text));
            else if (values[i] is byte[] bytes) writer.Write(bytes);
            else if (values[i] != null) throw new NotSupportedException("Probe variable type is not implemented.");
        }
        if (offsets.Count != 0) { writer.Write(checked((short)data.Position)); for (int i = offsets.Count - 1; i >= 0; i--) writer.Write(offsets[i]); writer.Write((short)offsets.Count); }
        writer.Write(mask); writer.Flush(); return data.ToArray();
    }

    private static byte[] IndexPage(Table table, Index index) {
        var entries = new List<byte[]>();
        // ASCII weights observed in independently produced catalog-name keys. This probe rejects other names.
        var weights = new Dictionary<char, byte> { { 'a', 0x4a }, { 'b', 0x4c }, { 'c', 0x4d }, { 'd', 0x4f }, { 'e', 0x51 }, { 'g', 0x55 }, { 'h', 0x57 }, { 'i', 0x59 }, { 'j', 0x5b }, { 'l', 0x5e }, { 'm', 0x60 }, { 'n', 0x62 }, { 'o', 0x64 }, { 'p', 0x66 }, { 'q', 0x68 }, { 'r', 0x69 }, { 's', 0x6b }, { 't', 0x6d }, { 'u', 0x6f }, { 'y', 0x76 } };
        for (int row = 0; row < table.Rows.Length; row++) {
            using var entry = new MemoryStream();
            foreach (int column in index.Columns) {
                object? value = table.Rows[row][column];
                if (value == null) { entry.WriteByte(0); continue; }
                entry.WriteByte(0x7f);
                if (value is string text) { foreach (char ch in text.ToLowerInvariant()) { if (!weights.TryGetValue(ch, out byte weight)) throw new NotSupportedException("Probe collation does not cover this character."); entry.WriteByte(weight); } entry.WriteByte(1); entry.WriteByte(0); } else { var integer = new byte[4]; BinaryPrimitives.WriteUInt32BigEndian(integer, unchecked((uint)Convert.ToInt32(value)) ^ 0x80000000); entry.Write(integer); }
            }
            entry.WriteByte((byte)(table.DataPage >> 16)); entry.WriteByte((byte)(table.DataPage >> 8)); entry.WriteByte((byte)table.DataPage); entry.WriteByte((byte)row); entries.Add(entry.ToArray());
        }
        entries.Sort((a, b) => { for (int i = 0; i < Math.Min(a.Length, b.Length); i++) { int result = a[i].CompareTo(b[i]); if (result != 0) return result; } return a.Length.CompareTo(b.Length); });
        var page = new byte[PageSize]; page[0] = 4; page[1] = 1; U32(page, 4, (uint)table.DefinitionPage); int end = 480;
        foreach (byte[] entry in entries) { if (end + entry.Length > PageSize) throw new InvalidDataException("Probe index exceeds one leaf page."); entry.CopyTo(page, end); end += entry.Length; int bit = end - 480; page[27 + bit / 8] |= (byte)(1 << (bit % 8)); }
        U16(page, 2, (ushort)(PageSize - end)); return page;
    }

    private static byte[] UsageMap(IEnumerable<int> pages) {
        var map = new byte[69];
        foreach (int page in pages) map[5 + page / 8] |= (byte)(1 << (page % 8));
        return map;
    }

    private static byte[] DataPage(int owner, byte[][] rows) {
        var page = new byte[PageSize]; page[0] = 1; page[1] = 1; U32(page, 4, (uint)owner); U16(page, 12, checked((ushort)rows.Length));
        int end = PageSize;
        for (int i = 0; i < rows.Length; i++) { end -= rows[i].Length; if (end < 14 + rows.Length * 2) throw new InvalidDataException("Probe rows exceed one page."); rows[i].CopyTo(page, end); U16(page, 14 + i * 2, (ushort)end); }
        U16(page, 2, (ushort)(end - 14 - rows.Length * 2)); return page;
    }
    private static void U16(byte[] bytes, int offset, ushort value) => BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset, 2), value);
    private static void U32(byte[] bytes, int offset, uint value) => BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset, 4), value);
}