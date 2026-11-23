namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Owns an enumerated source until projection. Chunks avoid copying a growing
        /// reference array, and projected rows can be released while selector values accumulate.
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

        /// <summary>Writes rows through typed columns along the selector route, without boxing their values.</summary>
        private void InsertTypedObjects<T>(IEnumerable<T> items, ExcelDataColumns<T> columns, bool includeHeaders, int startRow) {
            ArgumentNullException.ThrowIfNull(items);
            if (columns.Count == 0) {
                throw new ArgumentException("At least one column is required.", nameof(columns));
            }

            ValidateObjectExportDenseDimensions(rowCount: 0, columns.Count, includeHeaders: false);
            IReadOnlyList<T> rows = MaterializeObjectRowsBounded(items, includeHeaders, startRow);
            if (rows.Count == 0) {
                return;
            }

            ValidateObjectExportDenseDimensions(rows.Count, columns.Count, includeHeaders);
            var headers = new string[columns.Count];
            bool hasBlankDisplayHeader = false;
            for (int c = 0; c < headers.Length; c++) {
                string header = columns[c].Header ?? $"Column{c + 1}";
                headers[c] = header;
                if (includeHeaders && string.IsNullOrWhiteSpace(header)) hasBlankDisplayHeader = true;
            }

            var values = new ExcelTabularCellValues(headers.Length, rows.Count);
            var inferredColumnTypes = new Type?[headers.Length];
            var snapshot = rows as ObjectExportRowSnapshot<T>;
            int index = 0;
            for (int r = 0; r < rows.Count; r++) {
                T item = rows[r];
                for (int c = 0; c < headers.Length; c++) {
                    UpdateObjectExportColumnType(inferredColumnTypes, c, columns[c].Store(item, values, index++));
                }
                snapshot?.Release(r);
            }

            WriteProjectedObjectValues(headers, values, inferredColumnTypes, includeHeaders, startRow, hasBlankDisplayHeader);
        }

        /// <summary>Projects an owned snapshot after enumeration, then releases each consumed row.</summary>
        private static void FillOwnedObjectExportValues<T>(
            ObjectExportRowSnapshot<T> rows,
            Func<T, object?>[] selectors,
            ExcelTabularCellValues values,
            Type?[] inferredColumnTypes) {
            int index = 0;
            for (int row = 0; row < rows.Count; row++) {
                T item = rows[row];
                for (int column = 0; column < selectors.Length; column++) {
                    object? value = selectors[column](item);
                    values.Set(index++, value);
                    UpdateObjectExportColumnType(inferredColumnTypes, column, value);
                }
                rows.Release(row);
            }
        }
    }
}
