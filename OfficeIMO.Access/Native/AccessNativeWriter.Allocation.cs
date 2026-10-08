namespace OfficeIMO.Access;

internal sealed partial class AccessNativeWriter {
    private void WriteMaps(Table table) {
        var maps = new List<byte[]> { UsageMap(table.DataPages), UsageMap(table.DataPages.Where(p => _pages[p][2] != 0 || _pages[p][3] != 0)) };
        foreach (Index index in PhysicalIndexes(table)) maps.Add(UsageMap(index.Pages));
        foreach (Column column in table.Columns.Where(c => c.Type == 11 || c.Type == 12)) { maps.Add(UsageMap(column.LongPages)); maps.Add(UsageMap(Array.Empty<int>())); }
        var rows = new List<byte[]>(); int number = table.MapPage, size = 14;
        foreach (byte[] map in maps) {
            if (size + map.Length + 2 > PageSize || rows.Count == 255) {
                _pages[number] = DataPage(0, rows); number = Allocate(); rows.Clear(); size = 14;
            }
            table.MapPointers.Add((uint)((number << 8) | rows.Count)); rows.Add(map); size += map.Length + 2;
        }
        _pages[number] = DataPage(0, rows);
    }
    private byte[] UsageMap(IEnumerable<int> ownedPages) {
        int[] owned = ownedPages.ToArray(); int maximum = owned.Length == 0 ? 0 : owned.Max();
        if (maximum < 512) {
            var map = new byte[69]; foreach (int page in owned) map[5 + page / 8] |= (byte)(1 << (page % 8)); return map;
        }
        var references = new byte[69]; references[0] = 1;
        foreach (var group in owned.GroupBy(p => p / 32736)) {
            int number = Allocate(); byte[] page = _pages[number]; page[0] = 5; page[1] = 1;
            foreach (int bit in group.Select(p => p % 32736)) page[4 + bit / 8] |= (byte)(1 << (bit % 8));
            U32(references, 1 + group.Key * 4, (uint)number);
        }
        return references;
    }
    private void WriteGlobalMap() {
        // Every emitted page is allocated. Beyond the bounded plan, pages are available to the native engine.
        if (_pages.Count < 512) {
            byte[] free = UsageMap(Enumerable.Range(0, _pages.Count));
            for (int i = 5; i < free.Length; i++) free[i] = (byte)~free[i];
            _pages[1] = DataPage(1, new[] { free, free }); return;
        }
        int mapCount = (_pages.Count + 1) / 32736 + 1;
        var numbers = new int[mapCount]; for (int i = 0; i < mapCount; i++) numbers[i] = Allocate();
        var references = new byte[69]; references[0] = 1;
        for (int i = 0; i < mapCount; i++) {
            byte[] page = _pages[numbers[i]]; page[0] = 5; page[1] = 1;
            for (int p = Math.Max(_pages.Count, i * 32736); p < (i + 1) * 32736; p++) { int bit = p % 32736; page[4 + bit / 8] |= (byte)(1 << (bit % 8)); }
            U32(references, 1 + i * 4, (uint)numbers[i]);
        }
        _pages[1] = DataPage(1, new[] { references, references });
    }
}
