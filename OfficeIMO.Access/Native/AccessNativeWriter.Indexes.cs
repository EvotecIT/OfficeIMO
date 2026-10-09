namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        private static readonly byte[] LetterWeights = { 0x4a, 0x4c, 0x4d, 0x4f, 0x51, 0x53, 0x55, 0x57, 0x59, 0x5b, 0x5c, 0x5e, 0x60, 0x62, 0x64, 0x66, 0x68, 0x69, 0x6b, 0x6d, 0x6f, 0x71, 0x73, 0x75, 0x76, 0x78 };
        private static byte[] TextKey(string value) {
            using MemoryStream bytes = new MemoryStream();
            bool directory = value.StartsWith("\u0003", StringComparison.Ordinal);
            if (directory) value = value.Substring(1);
            // General legacy (1033) weights observed in independently produced Jet4 and ACE12 keys.
            foreach (char c in value.TrimEnd(' ')) {
                if (c >= 'a' && c <= 'z') bytes.WriteByte(LetterWeights[c - 'a']);
                else if (c >= 'A' && c <= 'Z') bytes.WriteByte(LetterWeights[c - 'A']);
                else if (c >= '0' && c <= '9') bytes.WriteByte((byte)(0x36 + (c - '0') * 2));
                else if (c == '_') { bytes.WriteByte(0x2b); bytes.WriteByte(3); }
                else if (c == ' ') bytes.WriteByte(7);
                else throw new NotSupportedException("Native text keys currently qualify ASCII letters, digits, underscore and spaces. Other indexed text and catalog names require additional collation qualification.");
            }
            bytes.WriteByte(1);
            if (directory) {
                // Qualified leading directory-name control character: General legacy's unprintable weight.
                bytes.WriteByte(1); bytes.WriteByte(1); bytes.WriteByte(1);
                bytes.WriteByte(0x80); bytes.WriteByte(7); bytes.WriteByte(6); bytes.WriteByte(5);
            }
            bytes.WriteByte(0); return bytes.ToArray();
        }
        private static byte[] Key(Table table, Index index, object?[] row) {
            using MemoryStream key = new MemoryStream();
            foreach (int ordinal in index.Columns) {
                object? value = row[ordinal]; if (value == null) { key.WriteByte(0); continue; }
                key.WriteByte(0x7f); Column column = table.Columns[ordinal]; byte[] bytes;
                switch (column.Type) {
                    case 2: bytes = new[] { (byte)value }; break;
                    case 3: bytes = new byte[2]; BigEndian(bytes, 0, (ushort)(unchecked((ushort)(short)value) ^ 0x8000), 2); break;
                    case 4: bytes = new byte[4]; BigEndian(bytes, 0, unchecked((uint)(int)value) ^ 0x80000000); break;
                    case 10: bytes = TextKey((string)value); break;
                    default: throw new NotSupportedException("Native index creation currently qualifies Byte, Int16, Int32/AutoNumber and bounded text keys.");
                }
                key.Write(bytes, 0, bytes.Length);
            }
            if (key.Length > 510) throw new NotSupportedException("Native creation does not truncate index keys longer than 510 bytes.");
            return key.ToArray();
        }
        private void WriteIndex(Table table, Index index) {
            List<byte[]> entries = new List<byte[]>(); HashSet<string> keys = new HashSet<string>();
            for (int r = 0; r < table.Rows.Length; r++) {
                _cancellation.ThrowIfCancellationRequested(); object?[] row = table.Rows[r];
                bool nullKey = index.Columns.All(c => row[c] == null);
                if ((index.Flags & 8) != 0 && index.Columns.Any(c => row[c] == null)) throw new InvalidDataException("A native primary or required index cannot contain null fields.");
                if ((index.Flags & 2) != 0 && nullKey) continue;
                byte[] key = Key(table, index, row);
                if (!keys.Add(Convert.ToBase64String(key)) && (index.Flags & 1) != 0 && !nullKey) throw new InvalidDataException("Modeled rows violate a native unique index.");
                byte[] entry = new byte[key.Length + 4]; Buffer.BlockCopy(key, 0, entry, 0, key.Length);
                uint id = table.RowIds[r]; entry[key.Length] = (byte)(id >> 24); entry[key.Length + 1] = (byte)(id >> 16); entry[key.Length + 2] = (byte)(id >> 8); entry[key.Length + 3] = (byte)id;
                entries.Add(entry);
            }
            index.UniqueCount = keys.Count; entries.Sort(CompareBytes);
            WriteIndexEntries(table, index, entries);
        }
        private void WriteIndexEntries(Table table, Index index, List<byte[]> entries) {
            List<List<byte[]>> groups = PackIndexEntries(entries);
            if (groups.Count == 1) { _pages[index.RootPage] = IndexPage(table.DefinitionPage, groups[0]); index.Pages.Add(index.RootPage); return; }
            List<(int Page, byte[] Last)> layer = new List<(int Page, byte[] Last)>();
            for (int g = 0; g < groups.Count; g++) { int number = Allocate(); index.Pages.Add(number); layer.Add((number, groups[g][groups[g].Count - 1])); }
            for (int g = 0; g < groups.Count; g++) _pages[layer[g].Page] = IndexPage(table.DefinitionPage, groups[g], previous: g == 0 ? 0 : layer[g - 1].Page, next: g + 1 < layer.Count ? layer[g + 1].Page : 0);
            byte level = 1;
            while (true) {
                _cancellation.ThrowIfCancellationRequested();
                List<byte[]> branches = new List<byte[]>();
                for (int i = 0; i < layer.Count - 1; i++) { byte[] entry = new byte[layer[i].Last.Length + 4]; Buffer.BlockCopy(layer[i].Last, 0, entry, 0, layer[i].Last.Length); BigEndian(entry, entry.Length - 4, (uint)layer[i].Page); branches.Add(entry); }
                if (branches.Sum(x => x.Length) <= PageSize - 480) {
                    _pages[index.RootPage] = IndexPage(table.DefinitionPage, branches, level, tail: layer[layer.Count - 1].Page); index.Pages.Add(index.RootPage); return;
                }
                List<(int Page, byte[] Last)> nextLayer = new List<(int Page, byte[] Last)>(); int position = 0;
                while (position < layer.Count) {
                    int begin = position, size = 0; List<byte[]> group = new List<byte[]>();
                    while (position < layer.Count - 1 && size + branches[position].Length <= PageSize - 480) { group.Add(branches[position]); size += branches[position].Length; position++; }
                    int tail = layer[position].Page; byte[] last = layer[position].Last; position++;
                    int page = Allocate(); index.Pages.Add(page); _pages[page] = IndexPage(table.DefinitionPage, group, level, tail: tail); nextLayer.Add((page, last));
                    if (position == begin) throw new InvalidDataException("Native index construction makes no progress.");
                }
                layer = nextLayer; level = checked((byte)(level + 1));
            }
        }
        private static List<List<byte[]>> PackIndexEntries(List<byte[]> entries) {
            List<List<byte[]>> groups = new List<List<byte[]>> { new List<byte[]>() }; int size = 0;
            foreach (byte[] entry in entries) {
                if (size + entry.Length > PageSize - 480) { groups.Add(new List<byte[]>()); size = 0; }
                groups[groups.Count - 1].Add(entry); size += entry.Length;
            }
            return groups;
        }
        private static byte[] IndexPage(int owner, List<byte[]> entries, byte level = 0, int previous = 0, int next = 0, int tail = 0) {
            byte[] page = new byte[PageSize]; page[0] = level == 0 ? (byte)4 : (byte)3; page[1] = 1; U32(page, 4, (uint)owner);
            U32(page, 12, (uint)previous); U32(page, 16, (uint)next); U32(page, 20, (uint)tail); page[26] = level;
            int end = 480;
            foreach (byte[] entry in entries) { Buffer.BlockCopy(entry, 0, page, end, entry.Length); end += entry.Length; int bit = end - 480; page[27 + bit / 8] |= (byte)(1 << (bit % 8)); }
            U16(page, 2, (ushort)(PageSize - end)); return page;
        }
        private static int CompareBytes(byte[] a, byte[] b) {
            for (int i = 0; i < Math.Min(a.Length, b.Length); i++) { int result = a[i].CompareTo(b[i]); if (result != 0) return result; } return a.Length.CompareTo(b.Length);
        }
    }
}
