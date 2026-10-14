using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        /// <summary>Extends the native free-page maps while retaining the source's free-page state.</summary>
        private void WriteMutationGlobalMaps(AccessNativeDatabase database) {
            var original = database.Page(1, 1);
            if (I32(original, 4) != 1 || AccessNativeBinary.U16(original, 12) != 2)
                throw new NotSupportedException("The native global allocation map layout is not qualified for mutation.");
            bool[][] free = { ReadGlobalFreePages(database, 0), ReadGlobalFreePages(database, 1) };
            if (_pages.Count < 512) {
                byte[][] rows = new byte[2][];
                for (int row = 0; row < 2; row++) {
                    rows[row] = new byte[69];
                    for (int page = 0; page < 512; page++)
                        if (page >= _pages.Count || page < database.PageCount && free[row][page])
                            rows[row][5 + page / 8] |= (byte)(1 << (page % 8));
                }
                _pages[1] = DataPage(1, rows); return;
            }
            // Reserve both reference-map sets before deciding which pages are available.
            int groups = checked((_pages.Count + 32735) / 32736);
            while ((long)_pages.Count + groups * 2 > (long)groups * 32736) groups++;
            if (groups > 16) throw new InvalidDataException("Native allocation maps exceed the qualified output capacity.");
            int[][] numbers = { new int[groups], new int[groups] };
            for (int row = 0; row < 2; row++) for (int group = 0; group < groups; group++) numbers[row][group] = Allocate();
            byte[][] references = { new byte[69], new byte[69] };
            for (int row = 0; row < 2; row++) {
                references[row][0] = 1;
                for (int group = 0; group < groups; group++) {
                    _cancellation.ThrowIfCancellationRequested(); byte[] page = _pages[numbers[row][group]]; page[0] = 5; page[1] = 1;
                    for (int bit = 0; bit < 32736; bit++) {
                        int number = group * 32736 + bit;
                        if (number >= _pages.Count || number < database.PageCount && free[row][number])
                            page[4 + bit / 8] |= (byte)(1 << (bit % 8));
                    }
                    U32(references[row], 1 + group * 4, (uint)numbers[row][group]);
                }
            }
            _pages[1] = DataPage(1, references);
        }

        private bool[] ReadGlobalFreePages(AccessNativeDatabase database, int row) {
            var map = database.Row(1, row, false, _cancellation);
            if (map.Length < 5) throw new InvalidDataException("The native global free-page map is truncated.");
            bool[] free = new bool[database.PageCount];
            if (map[0] == 0) {
                uint start = AccessNativeBinary.U32(map, 1);
                for (int page = 0; page < free.Length; page++) {
                    long bit = page - (long)start;
                    // An inline global map treats pages outside its represented range as free.
                    free[page] = bit < 0 || bit >= (map.Length - 5L) * 8 || (map[5 + (int)bit / 8] & (1 << ((int)bit % 8))) != 0;
                }
            } else if (map[0] == 1) {
                if ((map.Length - 1) % 4 != 0) throw new InvalidDataException("The native global reference map is malformed.");
                HashSet<int> seen = new HashSet<int>();
                for (int group = 0; group * 32736L < free.Length; group++) {
                    _cancellation.ThrowIfCancellationRequested(); int offset = 1 + group * 4;
                    int reference = offset <= map.Length - 4 ? I32(map, offset) : 0;
                    if (reference != 0 && !seen.Add(reference)) throw new InvalidDataException("Native global allocation maps repeat a reference page.");
                    var bits = reference == 0 ? default : database.Page(reference, 5);
                    for (int bit = 0; bit < Math.Min(32736, free.Length - group * 32736); bit++)
                        free[group * 32736 + bit] = reference == 0 || (bits[4 + bit / 8] & (1 << (bit % 8))) != 0;
                }
            } else throw new NotSupportedException("The native global allocation-map type is not qualified.");
            if (free[0] || free[1]) throw new InvalidDataException("Native allocation maps mark the header or global map as free.");
            return free;
        }
    }
}
