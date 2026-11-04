namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Owns an enumerated source until projection. Chunks avoid copying a growing
        /// reference array, and projected rows can be released while selector boxes accumulate.
        /// Caller-owned indexed sources retain their ordinary per-selector indexing path.
        /// </summary>
        private sealed class ObjectExportRowSnapshot<T> : IReadOnlyList<T> {
            private readonly List<T[]> _chunks = new List<T[]>();
            private readonly int _chunkBits;
            private readonly int _chunkMask;
            private readonly int _maximumRows;

            internal ObjectExportRowSnapshot(int capacity, int maximumRows) {
                _chunkBits = 12;
                if (capacity > 0 && capacity < (1 << _chunkBits)) {
                    _chunkBits = 0;
                    while ((1 << _chunkBits) < capacity) _chunkBits++;
                }
                _chunkMask = (1 << _chunkBits) - 1;
                _maximumRows = maximumRows;
            }

            public int Count { get; private set; }

            public T this[int index] {
                get {
                    if ((uint)index >= (uint)Count) throw new ArgumentOutOfRangeException(nameof(index));
                    return _chunks[index >> _chunkBits][index & _chunkMask];
                }
            }

            internal void Add(T value) {
                int offset = Count & _chunkMask;
                if (offset == 0) {
                    _chunks.Add(new T[Math.Min(_chunkMask + 1, _maximumRows - Count)]);
                }
                _chunks[Count >> _chunkBits][offset] = value;
                Count++;
            }

            internal void Release(int index) {
                int chunk = index >> _chunkBits;
                _chunks[chunk][index & _chunkMask] = default!;
                if ((index & _chunkMask) == _chunkMask || index == Count - 1) {
                    _chunks[chunk] = Array.Empty<T>();
                }
            }

            public IEnumerator<T> GetEnumerator() {
                for (int index = 0; index < Count; index++) yield return this[index];
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }

        /// <summary>Projects an owned snapshot after enumeration, then releases each consumed row.</summary>
        private static void FillOwnedObjectExportValues<T>(
            ObjectExportRowSnapshot<T> rows,
            Func<T, object?>[] selectors,
            object?[] values,
            Type?[] inferredColumnTypes) {
            int index = 0;
            for (int row = 0; row < rows.Count; row++) {
                T item = rows[row];
                for (int column = 0; column < selectors.Length; column++) {
                    object? value = selectors[column](item);
                    values[index++] = value;
                    UpdateObjectExportColumnType(inferredColumnTypes, column, value);
                }
                rows.Release(row);
            }
        }
    }
}
